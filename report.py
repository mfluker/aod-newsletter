import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
from urllib.parse import quote_plus
from io import StringIO
from dataclasses import dataclass

import pandas as pd
import pytz
import requests
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer


@dataclass
class DateRange:
    """Represents a date range with formatting utilities."""
    start: datetime
    end: datetime
    
    def format(self, fmt: str = "%m/%d/%Y") -> Tuple[str, str]:
        """Format start and end dates."""
        return self.start.strftime(fmt), self.end.strftime(fmt)
    
    def get_last_year(self) -> 'DateRange':
        """Get the same date range from last year."""
        return DateRange(
            start=self.start.replace(year=self.start.year - 1),
            end=self.end.replace(year=self.end.year - 1)
        )


class CanvasSession:
    """Manages Canvas authentication and session handling."""
    
    COOKIE_PATH = "canvas_cookies.json"
    BASE_URL = "https://canvas.artofdrawers.com"
    
    def __init__(self):
        self.session = self._create_session()
    
    def _create_session(self) -> requests.Session:
        """Create authenticated session with Canvas cookies."""
        try:
            with open(self.COOKIE_PATH, "r") as f:
                raw_cookies = json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"Cookie file not found: {self.COOKIE_PATH}")
        
        session = requests.Session()
        for cookie in raw_cookies:
            params = {
                key: cookie[key] for key in ["domain", "path", "secure"] 
                if key in cookie
            }
            if "expirationDate" in cookie:
                params["expires"] = int(cookie["expirationDate"])
            
            session.cookies.set(
                cookie["name"], 
                cookie["value"], 
                **params
            )
        
        return session
    
    def get(self, url: str, **kwargs) -> requests.Response:
        """Make authenticated GET request."""
        response = self.session.get(url, **kwargs)
        response.raise_for_status()
        return response
    
    def post(self, url: str, **kwargs) -> requests.Response:
        """Make authenticated POST request."""
        response = self.session.post(url, **kwargs)
        response.raise_for_status()
        return response


class JobsStatusScraper:
    """Handles scraping job status data from Canvas."""
    
    STATUS_FILTERS = {
        "Submitted to Manufacturing Partner": 5,
        "Order Shipped": 6,
    }
    
    URL_TEMPLATE = """
    https://canvas.artofdrawers.com/listjobs.html?dsraas=1&id=&location_id=&zone=&zone_id=&production_priority_ge=&production_priority_le=&opportunity=&opportunity_id=&customer=&customer_id=&campaign_source=&customer_id_sub_filters_campaign_source_id=&customer_id_sub_filters_firstname=&customer_id_sub_filters_lastname=&customer_id_sub_filters_spouse=&customer_id_sub_filters_preferred_phone=&customer_id_sub_filters_cell_phone=&customer_id_sub_filters_emailaddr=&city=&state_id=&country_id=&latitude_ge=&latitude_le=&longitude_ge=&longitude_le=&location_tax_rate_id=&total_cost_ge=&total_cost_le=&material_total_ge=&material_total_le=&labor_total_ge=&labor_total_le=&delivery_total_ge=&delivery_total_le=&discount_total_ge=&discount_total_le=&credit_memo_total_ge=&credit_memo_total_le=&tax_total_ge=&tax_total_le=&order_total_ge=&order_total_le=&amount_paid_ge=&amount_paid_le=&amount_due_ge=&amount_due_le=&designer_id=&tma_id=&relationship_partner_id=&installer_id=&shipping_type_id=&number_of_items_ge=&number_of_items_le=&manufacturing_batch_id=&manufacturing_facility_id=&manufacturing_status_id=&date_submitted_to_manufacturing_ge=&date_submitted_to_manufacturing_le=&date_submitted_to_manufacturing_r=select&number_of_days_ago_submitted_to_go_ge=&number_of_days_ago_submitted_to_go_le=&number_of_biz_days_at_manufacturing_status_ge=&number_of_biz_days_at_manufacturing_status_le=&date_submitted_to_manufacturing_partner_ge=&date_submitted_to_manufacturing_partner_le=&date_submitted_to_manufacturing_partner_r=select&date_projected_to_ship_ge=&date_projected_to_ship_le=&date_projected_to_ship_r=select&date_shipped_ge=&date_shipped_le=&date_shipped_r=select&carrier_id=&tracking_number=&date_delivered_ge=&date_delivered_le=&date_delivered_r=select&commission_rate_type_id=&designer_commission_override_percentage_ge=&designer_commission_override_percentage_le=&tma_commission_rate_type_id=&tma_commission_has_been_paid_y=y&tma_commission_has_been_paid_n=n&job_type_id=&current_status_ids%5B%5D=2&current_status_ids%5B%5D=3&current_status_ids%5B%5D=5&current_status_ids%5B%5D=6&current_status_ids%5B%5D=7&current_status_ids%5B%5D=8&current_status_ids%5B%5D=9&current_status_ids%5B%5D=10&current_status_ids%5B%5D=11&current_status_ids%5B%5D=12&current_status_ids%5B%5D=13&current_status_ids%5B%5D=21&current_status_ids%5B%5D=22&current_status_ids%5B%5D=23&current_status_ids%5B%5D=24&current_status_ids%5B%5D=25&current_status_ids%5B%5D=30&current_status_ids%5B%5D=31&current_status_ids%5B%5D=33&current_status_ids%5B%5D=34&current_status_ids%5B%5D=37&current_status_ids%5B%5D=38&date_of_last_status_change_ge=&date_of_last_status_change_le=&date_of_last_status_change_r=select&promotion_id=&date_placed_ge=&date_placed_le=&date_placed_r=select&date_of_initial_appointment_ge=&date_of_initial_appointment_le=&date_of_initial_appointment_r=select&date_of_welcome_call_ge=&date_of_welcome_call_le=&date_of_welcome_call_r=select&date_measurements_scheduled_ge=&date_measurements_scheduled_le=&date_measurements_scheduled_r=select&date_installation_scheduled_ge=&date_installation_scheduled_le=&date_installation_scheduled_r=select&date_of_final_payment_ge=&date_of_final_payment_le=&date_of_final_payment_r=select&date_completed_ge=&date_completed_le=&date_completed_r=select&date_last_payment_ge=&date_last_payment_le=&date_last_payment_r=select&payment_type_id=&memo=&payment_value_lookup=&time_est=&job_survey_response_id=&is_rush_y=y&is_rush_n=n&rush_is_billable_y=y&rush_is_billable_n=n&is_split_order_y=y&is_split_order_n=n&exclude_from_close_rate_y=y&exclude_from_close_rate_n=n&exclude_from_average_sale_y=y&exclude_from_average_sale_n=n&number_of_basics_ge=&number_of_basics_le=&number_of_classics_ge=&number_of_classics_le=&number_of_designers_ge=&number_of_designers_le=&number_of_shelves_ge=&number_of_shelves_le=&number_of_dividers_ge=&number_of_dividers_le=&number_of_accessories_ge=&number_of_accessories_le=&number_of_strip_mounts_ge=&number_of_strip_mounts_le=&number_of_other_ge=&number_of_other_le=&number_of_options_ge=&number_of_options_le=&nps_survey_rating_ge=&nps_survey_rating_le=&wm_note=&active_y=y&date_last_modified_ge=&date_last_modified_le=&date_last_modified_r=select&date_added_ge=&date_added_le=&date_added_r=select&status_field_name_for_filter=REPLACE_STATUS&status_update_search_date_ge=REPLACE_START&status_update_search_date_le=REPLACE_END&status_update_search_date_r=select&sort_by=id&sort_dir=DESC&display=on&c%5B%5D=id&c%5B%5D=location_id&filter=Submit
    """.strip()
    
    def __init__(self, session: CanvasSession):
        self.session = session
    
    def _build_url(self, status_filter_id: int, date_range: DateRange) -> str:
        """Build the jobs listing URL with parameters."""
        start_str, end_str = date_range.format()
        return (
            self.URL_TEMPLATE
            .replace("REPLACE_STATUS", str(status_filter_id))
            .replace("REPLACE_START", quote_plus(start_str))
            .replace("REPLACE_END", quote_plus(end_str))
        )
    
    def _classify_order_type(self, order_id: str) -> str:
        """Classify order type based on ID pattern."""
        if not isinstance(order_id, str):
            return "New"
        
        if order_id.startswith("C"):
            return "Claim"
        elif order_id.startswith("R"):
            return "Reorder"
        elif re.match(r"^\d", order_id):
            return "New"
        
        return "New"
    
    def _fetch_status_data(self, status_name: str, status_filter_id: int, date_range: DateRange) -> pd.DataFrame:
        """Fetch and parse job status data."""
        url = self._build_url(status_filter_id, date_range)
        response = self.session.get(url)
        
        # Clean HTML tags from response
        cleaned_text = re.sub(r"<[^>]+>", "", response.text).strip()
        if not cleaned_text:
            return pd.DataFrame()
        
        # Parse CSV data
        df = pd.read_csv(StringIO(cleaned_text), engine="python")
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        
        # Find date column
        date_column_name = f"{status_name} Date".lower()
        date_column = next(
            (col for col in df.columns if col.lower() == date_column_name), 
            None
        )
        
        # Add metadata columns
        df["Status"] = status_name
        df["Date"] = df[date_column] if date_column else pd.NaT
        df["Order Type"] = df["ID"].apply(self._classify_order_type)
        
        return df[["ID", "Order Type", "Franchisee", "Date", "Status"]]
    
    def count_jobs_by_status(self, status_name: str, date_range: DateRange) -> int:
        """Count jobs in a specific status for a date range."""
        if status_name not in self.STATUS_FILTERS:
            raise ValueError(f"Unknown status: {status_name}")
        
        status_filter_id = self.STATUS_FILTERS[status_name]
        df = self._fetch_status_data(status_name, status_filter_id, date_range)
        return len(df)
    
    def generate_combined_csv(self, date_range: DateRange, output_path: Optional[Path] = None) -> pd.DataFrame:
        """Generate combined CSV of all job statuses."""
        if output_path is None:
            start_str, end_str = date_range.format()
            filename = f"{start_str.replace('/', '')}_{end_str.replace('/', '')}_jobs.csv"
            output_path = Path("Reports") / filename
        
        dataframes = []
        for status_name, status_filter_id in self.STATUS_FILTERS.items():
            df = self._fetch_status_data(status_name, status_filter_id, date_range)
            if not df.empty:
                dataframes.append(df)
        
        if not dataframes:
            return pd.DataFrame()
        
        combined_df = pd.concat(dataframes, ignore_index=True)
        
        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        combined_df.to_csv(output_path, index=False)
        
        return combined_df


class ConversionReportDownloader:
    """Handles downloading and processing conversion reports."""
    
    def __init__(self, session: CanvasSession):
        self.session = session
        self.form_url = f"{session.BASE_URL}/scripts/lead-to-appointment-conversion/index.html"
        self.csv_url = f"{session.BASE_URL}/scripts/report_as_spreadsheet.html?report=report_lead_to_appointment_conversion"
    
    def download_report(self, date_range: DateRange) -> pd.DataFrame:
        """Download conversion report for specified date range."""
        start_str, end_str = date_range.format()
        
        # Submit form to set parameters
        payload = {
            "start_date": start_str,
            "end_date": end_str,
            "include_homeshow": "true",
            "quick_search": "Search",
            "search_for": "",
            "submit": "Show Report",
        }
        
        self.session.post(self.form_url, data=payload, headers={"Referer": self.form_url})
        
        # Download CSV data
        response = self.session.get(self.csv_url, headers={"Referer": self.form_url})
        csv_text = response.text
        
        # Validate response format
        if "Call Center Rep" not in csv_text.splitlines()[0]:
            raise ValueError(f"Unexpected response format from Canvas: {csv_text[:500]}")
        
        # Parse CSV
        try:
            df = pd.read_csv(StringIO(csv_text))
            df["Outbound Communication Count"] = df["Outbound Communication Count"].astype(int)
            return df
        except pd.errors.ParserError as e:
            raise ValueError(f"Failed to parse CSV: {csv_text[:500]}") from e
    
    def get_total_outbound_communications(self, date_range: DateRange) -> int:
        """Get total outbound communications for date range."""
        df = self.download_report(date_range)
        return df["Outbound Communication Count"].sum()
   

class PDFReportGenerator:
    """Generates PDF reports with year-over-year comparisons."""
    
    def __init__(self, output_dir: Path = Path("Reports")):
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def _format_yoy_stat(self, label: str, current: float, last_year: float, is_currency: bool = False) -> str:
        """Format year-over-year statistic with percentage change."""
        delta = current - last_year
        
        try:
            percent_change = (delta / last_year) * 100
        except ZeroDivisionError:
            percent_change = 0
        
        arrow = "↑" if percent_change >= 0 else "↓"
        percent_display = f"{abs(percent_change):.1f}%"
        
        if is_currency:
            current_display = f"${current:,.2f}"
            last_year_display = f"${last_year:,.2f}"
        else:
            current_display = f"{current:,}"
            last_year_display = f"{last_year:,}"
        
        return f"<b>{label}:</b> {current_display} ({arrow} {percent_display} vLY) ({last_year_display} LY)"

    def _format_duration_comparison(self, label: str, current_days: float, last_year_days: float) -> str:
        """
        Format an absolute comparison between two durations (in days).
        E.g. "24 days, 20 hours (↑ by 2 days) (22 days, 20 hours LY)"
        """
        # Human‐readable strings
        cur_days = int(current_days)
        cur_hours = int((current_days - cur_days) * 24)
        ly_days  = int(last_year_days)
        ly_hours = int((last_year_days  - ly_days ) * 24)
    
        # Difference & arrow
        diff_days = int(current_days - last_year_days)
        arrow     = "↑" if diff_days >= 0 else "↓"
        diff_abs  = abs(diff_days)
    
        # Build the final line
        return (
            f"<b>{label}:</b> "
            f"{cur_days} days, {cur_hours} hours "
            f"({arrow} by {diff_abs} days) "
            f"({ly_days} days, {ly_hours} hours LY)"
        )

    
    def create_report(self, data: Dict[str, Any], date_range: DateRange, pull_date: str) -> Path:
        """Create PDF report with provided data."""
        start_str, end_str = date_range.format()
        
        # Generate filename
        filename = f"AoD_Weekly_Newsletter_{pull_date.replace('/', '_')}.pdf"
        output_path = self.output_dir / filename
        
        # Prepare content
        styles = getSampleStyleSheet()
        story = []
        
        title = f"AoD Weekly Newsletter – {pull_date}"
        content_lines = [
            title,
            "",
            f"Data pulled on: {pull_date}",
            f"Period: {start_str} – {end_str}",
            "",
            self._format_yoy_stat("SSC Touches", data["ssc_current"], data["ssc_last_year"]),
            self._format_yoy_stat("Orders Shipped", data["shipped_current"], data["shipped_last_year"]),
            self._format_yoy_stat("Orders Submitted", data["submitted_current"], data["submitted_last_year"]),
            self._format_yoy_stat("Network Revenue", data["revenue_current"], data["revenue_last_year"], is_currency=True),
            # (
            #   f"<b>Avg. Time From Measurement to Shipped:</b> "
            #   f"{data['avg_meas_human']} "
            #   f"({data['avg_meas_diff_arrow']} by {data['avg_meas_diff_days']} days) "
            #   f"({data['avg_meas_ly_human']} LY)"
            # ),
            self._format_duration_comparison("Avg. Time From Measurement to Shipped", data["avg_meas_current"], data["avg_meas_lastyr"]),
            "",
        ]

        # if we got a top-3 DataFrame, append it
        if "top3_locations" in data and isinstance(data["top3_locations"], pd.DataFrame):
            top3 = data["top3_locations"]
            content_lines.append("")  # blank line
            content_lines.append("<b>Top 3 Locations by Revenue:</b>")
            for _, row in top3.iterrows():
                content_lines.append(f"{row['Rank']}. {row['Location']} – ${row['Revenue']:,.2f}")
        
        # Add content to PDF
        for line in content_lines:
            story.append(Paragraph(line, styles['Normal']))
            story.append(Spacer(1, 12))
        
        # Build PDF
        doc = SimpleDocTemplate(str(output_path), pagesize=LETTER)
        doc.build(story)
        
        return output_path


class NetworkRevenue:
    """Handles downloading and processing network revenue data."""
    
    def __init__(self, session: CanvasSession):
        self.session = session
        self.report_url = f"{session.BASE_URL}/scripts/location_sales_rankings.html"
    
    def _build_url(self, date_range: DateRange) -> str:
        """Build URL with date parameters for revenue report."""
        start_str, end_str = date_range.format()
        sd = quote_plus(start_str)
        ed = quote_plus(end_str)
        return f"{self.report_url}?sd={sd}&ed={ed}&presetdates=na"
    
    def _parse_revenue_table(self, html: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Parse revenue table from HTML response."""
        try:
            from bs4 import BeautifulSoup
        except ImportError:
            raise ImportError("BeautifulSoup4 is required for HTML parsing. Install with: pip install beautifulsoup4")
        
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find_all("table")[-1]  # Use the last table on the page
        
        rows = []
        total_row = None
        
        for tr in table.find_all("tr"):
            cols = tr.find_all("td")
            if len(cols) != 3:
                continue
            
            rank, name, revenue = [td.text.strip() for td in cols]
            
            # Clean revenue value
            revenue_clean = float(revenue.replace("$", "").replace(",", ""))
            
            if name.lower() == "total":
                total_row = {
                    "Rank": rank,
                    "Location": name,
                    "Revenue": revenue_clean
                }
            else:
                rows.append({
                    "Rank": int(rank),
                    "Location": name,
                    "Revenue": revenue_clean
                })
        
        if total_row is None:
            raise ValueError("Could not find total row in revenue table.")
        
        df_all = pd.DataFrame(rows)
        df_top3 = df_all.nlargest(3, "Revenue").copy() if not df_all.empty else pd.DataFrame()
        df_total = pd.DataFrame([total_row])
        
        return df_total, df_top3
    
    def get_revenue_data(self, date_range: DateRange) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Get network revenue data for specified date range."""
        url = self._build_url(date_range)
        response = self.session.get(url)
        
        return self._parse_revenue_table(response.text)
    
    def get_total_revenue(self, date_range: DateRange) -> float:
        """Get total network revenue for date range."""
        df_total, _ = self.get_revenue_data(date_range)
        return df_total.iloc[0]["Revenue"]
    
    def get_revenue_summary(self, date_range: DateRange) -> Tuple[float, pd.DataFrame]:
        """Get revenue summary with total and top 3 locations."""
        df_total, df_top3 = self.get_revenue_data(date_range)
        total_value = df_total.iloc[0]["Revenue"]
        return total_value, df_top3


class MeasurementShippedScraper:
    
    URL_TEMPLATE = """
    https://canvas.artofdrawers.com/listjobs.html?dsraas=1&id=&location_id=&location_id_sub_filters_exclude_from_reports_n=n&zone=&zone_id=&production_priority_ge=&production_priority_le=&opportunity=&opportunity_id=&customer=&customer_id=&campaign_source=&customer_id_sub_filters_campaign_source_id=&customer_id_sub_filters_firstname=&customer_id_sub_filters_lastname=&customer_id_sub_filters_spouse=&customer_id_sub_filters_preferred_phone=&customer_id_sub_filters_cell_phone=&customer_id_sub_filters_emailaddr=&city=&state_id=&country_id=&latitude_ge=&latitude_le=&longitude_ge=&longitude_le=&location_tax_rate_id=&total_cost_ge=&total_cost_le=&material_total_ge=&material_total_le=&labor_total_ge=&labor_total_le=&delivery_total_ge=&delivery_total_le=&discount_total_ge=&discount_total_le=&credit_memo_total_ge=&credit_memo_total_le=&tax_total_ge=&tax_total_le=&order_total_ge=&order_total_le=&amount_paid_ge=&amount_paid_le=&amount_due_ge=&amount_due_le=&siteuser=&designer_id=&tma_id=&relationship_partner_id=&installer_id=&shipping_type_id=&number_of_items_ge=&number_of_items_le=&manufacturing_batch_id=&manufacturing_facility_id=&manufacturing_status_id=&date_submitted_to_manufacturing_ge=&date_submitted_to_manufacturing_le=&date_submitted_to_manufacturing_r=select&number_of_days_ago_submitted_to_go_ge=&number_of_days_ago_submitted_to_go_le=&number_of_biz_days_at_manufacturing_status_ge=&number_of_biz_days_at_manufacturing_status_le=&date_submitted_to_manufacturing_partner_ge=&date_submitted_to_manufacturing_partner_le=&date_submitted_to_manufacturing_partner_r=select&date_projected_to_ship_ge=&date_projected_to_ship_le=&date_projected_to_ship_r=select&date_shipped_ge=REPLACE_START&date_shipped_le=REPLACE_END&date_shipped_r=select&carrier_id=&tracking_number=&date_delivered_ge=&date_delivered_le=&date_delivered_r=select&commission_rate_type_id=&designer_commission_override_percentage_ge=&designer_commission_override_percentage_le=&tma_commission_rate_type_id=&tma_commission_has_been_paid_y=y&tma_commission_has_been_paid_n=n&job_type_id=&%63urrent_status_ids%5B%5D=21&%63urrent_status_ids%5B%5D=22&%63urrent_status_ids%5B%5D=2&%63urrent_status_ids%5B%5D=3&%63urrent_status_ids%5B%5D=23&%63urrent_status_ids%5B%5D=5&%63urrent_status_ids%5B%5D=24&%63urrent_status_ids%5B%5D=33&%63urrent_status_ids%5B%5D=6&%63urrent_status_ids%5B%5D=34&%63urrent_status_ids%5B%5D=37&%63urrent_status_ids%5B%5D=38&%63urrent_status_ids%5B%5D=7&%63urrent_status_ids%5B%5D=8&%63urrent_status_ids%5B%5D=9&%63urrent_status_ids%5B%5D=30&%63urrent_status_ids%5B%5D=10&%63urrent_status_ids%5B%5D=11&%63urrent_status_ids%5B%5D=31&%63urrent_status_ids%5B%5D=25&%63urrent_status_ids%5B%5D=12&%63urrent_status_ids%5B%5D=13&date_of_last_status_change_ge=&date_of_last_status_change_le=&date_of_last_status_change_r=select&promotion_id=&date_placed_ge=&date_placed_le=&date_placed_r=select&date_of_initial_appointment_ge=&date_of_initial_appointment_le=&date_of_initial_appointment_r=select&date_of_welcome_call_ge=&date_of_welcome_call_le=&date_of_welcome_call_r=select&date_measurements_scheduled_ge=&date_measurements_scheduled_le=&date_measurements_scheduled_r=select&date_installation_scheduled_ge=&date_installation_scheduled_le=&date_installation_scheduled_r=select&date_of_final_payment_ge=&date_of_final_payment_le=&date_of_final_payment_r=select&date_completed_ge=&date_completed_le=&date_completed_r=select&date_last_payment_ge=&date_last_payment_le=&date_last_payment_r=select&payment_type_id=&memo=&payment_value_lookup=&time_est=&job_survey_response_id=&is_rush_y=y&is_rush_n=n&rush_is_billable_y=y&rush_is_billable_n=n&is_split_order_y=y&is_split_order_n=n&exclude_from_close_rate_y=y&exclude_from_close_rate_n=n&exclude_from_average_sale_y=y&exclude_from_average_sale_n=n&number_of_basics_ge=&number_of_basics_le=&number_of_classics_ge=&number_of_classics_le=&number_of_designers_ge=&number_of_designers_le=&number_of_shelves_ge=&number_of_shelves_le=&number_of_dividers_ge=&number_of_dividers_le=&number_of_accessories_ge=&number_of_accessories_le=&number_of_strip_mounts_ge=&number_of_strip_mounts_le=&number_of_other_ge=&number_of_other_le=&number_of_options_ge=&number_of_options_le=&nps_survey_rating_ge=&nps_survey_rating_le=&wm_note=&active_y=y&date_last_modified_ge=&date_last_modified_le=&date_last_modified_r=select&date_added_ge=&date_added_le=&date_added_r=select&status_field_name_for_filter=23&status_update_search_date_ge=&status_update_search_date_le=&status_update_search_date_r=inpast&sort_by=id&sort_dir=DESC&c%5B%5D=id&c%5B%5D=location_id&c%5B%5D=order_total&c%5B%5D=date_shipped&filter=Submit
    """.strip()
    
    def __init__(self, session: CanvasSession):
        self.session = session
    
    def _build_url(self, date_range: DateRange) -> str:
        start_str, end_str = date_range.format()
        return (
            self.URL_TEMPLATE
            .replace("REPLACE_START", quote_plus(start_str))
            .replace("REPLACE_END",   quote_plus(end_str))
        )
    
    # def measurement_to_shipped(self, date_range: DateRange) -> str:
    #     """Return average “measurement → shipped” as “X days, Y hours”."""
    #     # use the session you stored on self
    #     resp = self.session.get(self._build_url(date_range))
    #     resp.raise_for_status()
        
    #     # strip tags & read CSV
    #     cleaned = re.sub(r"<[^>]+>", "", resp.text).strip()
    #     if not cleaned:
    #         return "No data returned"
    #     df = pd.read_csv(StringIO(cleaned), engine="python")
    #     df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    #     df["Date Shipped"] = pd.to_datetime(df["Date Shipped"], errors="coerce")
    #     def parse_last_measurement(s: str) -> pd.Timestamp:
    #         if pd.isna(s):
    #             return pd.NaT
    #         last = s.split(",")[-1].strip()
    #         return pd.to_datetime(last, errors="coerce")
    #     df["Measurement Approved Date"] = df["Measurement Approved Date"].apply(parse_last_measurement)

    #     df["Days Difference"] = (
    #         (df["Date Shipped"] - df["Measurement Approved Date"])
    #         .dt.total_seconds()
    #         .div(86400)
    #     )
    #     valid = df["Days Difference"].dropna()
    #     if valid.empty:
    #         return "No valid date pairs to average"

    #     avg = valid.mean()
    #     days = int(avg)
    #     hours = int((avg - days) * 24)
    #     return f"{days} days, {hours} hours"

    def measurement_to_shipped(self, date_range: DateRange) -> Tuple[float,str]:
        """
        Returns (avg_days_float, human_str) where human_str is
        'X days, Y hours'.
        """
        resp = self.session.get(self._build_url(date_range))
        resp.raise_for_status()
        cleaned = re.sub(r"<[^>]+>", "", resp.text).strip()
        if not cleaned:
            return 0.0, "No data"
        df = pd.read_csv(StringIO(cleaned), engine="python")
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    
        df["Date Shipped"] = pd.to_datetime(df["Date Shipped"], errors="coerce")
        df["Measurement Approved Date"] = (
            df["Measurement Approved Date"]
              .apply(lambda s: pd.to_datetime(s.split(",")[-1].strip(), errors="coerce"))
        )
    
        diffs = (
            (df["Date Shipped"] - df["Measurement Approved Date"])
            .dt.total_seconds()
            .div(86400)
            .dropna()
        )
        if diffs.empty:
            return 0.0, "No valid dates"
    
        avg = diffs.mean()
        days = int(avg)
        hours = int((avg - days) * 24)
        human = f"{days} days, {hours} hours"
        
        return avg, human



class WeeklyReportGenerator:
    """Main orchestrator for generating weekly reports."""
    
    def __init__(self):
        self.session = CanvasSession()
        self.jobs_scraper = JobsStatusScraper(self.session)
        self.conversion_downloader = ConversionReportDownloader(self.session)
        self.network_revenue = NetworkRevenue(self.session)
        self.meas_shipped_scraper  = MeasurementShippedScraper(self.session)
        self.pdf_generator = PDFReportGenerator()
    
    def generate_report(self, days_back: int = 30) -> Path:
        """Generate complete weekly report."""
        # Setup date ranges
        eastern = pytz.timezone("US/Eastern")
        today = datetime.now(eastern).date()
        
        current_range = DateRange(
            start=today - timedelta(days=days_back),
            end=today
        )
        last_year_range = current_range.get_last_year()
        
        pull_date_str = today.strftime("%m/%d/%Y")
        
        print(f"Generating report for {current_range.format()[0]} to {current_range.format()[1]}")

        # 1) revenue
        revenue_current, top3_df   = self.network_revenue.get_revenue_summary(current_range)
        revenue_last_year, _       = self.network_revenue.get_revenue_summary(last_year_range)

        # 2) measurement→shipped
        avg_cur, avg_cur_str       = self.meas_shipped_scraper.measurement_to_shipped(current_range)
        avg_ly,  avg_ly_str        = self.meas_shipped_scraper.measurement_to_shipped(last_year_range)

        # Collect all data
        data = {
            # SSC Touches
            "ssc_current": self.conversion_downloader.get_total_outbound_communications(current_range),
            "ssc_last_year": self.conversion_downloader.get_total_outbound_communications(last_year_range),
            
            # Orders Shipped
            "shipped_current": self.jobs_scraper.count_jobs_by_status("Order Shipped", current_range),
            "shipped_last_year": self.jobs_scraper.count_jobs_by_status("Order Shipped", last_year_range),
            
            # Orders Submitted
            "submitted_current": self.jobs_scraper.count_jobs_by_status("Submitted to Manufacturing Partner", current_range),
            "submitted_last_year": self.jobs_scraper.count_jobs_by_status("Submitted to Manufacturing Partner", last_year_range),

            # Network Revenue
            "revenue_current":   revenue_current,
            "revenue_last_year": revenue_last_year,

            # NEW metrics:
            "avg_meas_current": avg_cur,
            "avg_meas_lastyr":  avg_ly,

            # now include the top-3 locations DataFrame
            "top3_locations": top3_df,

        }
        
        # Generate PDF
        output_path = self.pdf_generator.create_report(data, current_range, pull_date_str)
        
        print(f"✅ PDF saved: {output_path.resolve()}")
        return output_path

def build_weekly_report(days_back: int = 30) -> Path:
    gen = WeeklyReportGenerator()
    return gen.generate_report(days_back)
