from datetime import datetime

import polars as pl

WOOCOMMERCE_PATH = "in/orders-2024-05-29-19-57-22.csv"
WPFORMS_PATH = "in/wpforms-3032-Artist-Registration-2024-05-29-20-00-19.xlsx"
YEAR = 2024
WEEK = 22


def filter_woocommerce(woocommerce, year, week):
    """Sets columns in WooCommerce report and filters for selected week."""
    return (
        woocommerce.select(
            pl.col("Order Date").alias("Date"),
            pl.col("First Name (Billing)").alias("First Name"),
            pl.col("Last Name (Billing)").alias("Last Name"),
            pl.col("Item Name").alias("Product(s)"),
            pl.col("Coupon Code").alias("Coupon(s)"),
            pl.col("Order Total Amount").alias("N. Revenue"),
            pl.col("Email (Billing)").alias("email"),
        )
        .with_columns(
            pl.col("Date").str.to_datetime("%Y-%m-%d %H:%M"),
            pl.col("email").str.to_lowercase(),
        )
        .filter((pl.col("Date").dt.year() == year) & (pl.col("Date").dt.week() == week))
    )


def filter_wpforms(wpforms, year, week):
    """Sets columns in WPForms report and filters for selected week."""
    return (
        wpforms.select(
            pl.col("Entry Date").str.to_datetime("%B %d, %Y %I:%M %p").alias("date"),
            pl.col("Email").alias("email"),
            pl.col("Address").alias("address"),
            pl.col("City").alias("city"),
            pl.col("State").alias("state"),
            pl.col("Zip").alias("zip"),
            pl.col("Website Url").alias("website"),
            pl.col("Phone Number").alias("Phone"),
            pl.col("Name as it will appear in your map listing").alias("Name for Map"),
            pl.col("Medium as it will appear in your map listing").alias("Medium"),
            pl.col("During JPOS I will be showing at (check one)").alias(
                "Show space/location"
            ),
            pl.col("Business Name").alias("Name of Business"),
            (
                pl.coalesce(
                    [pl.col("if at home address"), pl.col("Business Address")]
                ).alias("Studio/Business Address")
            ),
        )
        .with_columns(pl.col("email").str.to_lowercase())
        .filter(
            (~pl.col("email").str.split("@").list.first().is_in(["ss", "cr", "xx"]))
            & (pl.col("date").dt.year() == year)
            & (pl.col("date").dt.week() == week)
        )
    )


def merge_reports(woocommerce, wpforms):
    """Merges Woocommerce and WPForms Reports."""
    return (
        woocommerce.join(wpforms, on="email", how="outer_coalesce")
        .select(
            "Date",
            "First Name",
            "Last Name",
            "Product(s)",
            "Coupon(s)",
            "N. Revenue",
            "email",
            "address",
            "city",
            "state",
            "zip",
            "website",
            "Phone",
            "Name for Map",
            "Medium",
            "Show space/location",
            "Name of Business",
            "Studio/Business Address",
        )
        .sort("Date", nulls_last=True)
        .unique(subset="email", keep="last")
        .sort("Date", nulls_last=True)
    )


def generate_report(woocommerce_path, wpforms_path, year, week):
    "Generates Registration Report."
    week_start = datetime.fromisocalendar(year=year, week=week, day=1).strftime(
        "%Y-%m-%d"
    )
    week_end = datetime.fromisocalendar(year=year, week=week, day=7).strftime(
        "%Y-%m-%d"
    )

    woocommerce = pl.read_csv(woocommerce_path)
    wpforms = pl.read_excel(wpforms_path)

    filtered_woocommerce = filter_woocommerce(woocommerce, year, week)
    filtered_wpforms = filter_wpforms(wpforms, year, week)

    report = merge_reports(filtered_woocommerce, filtered_wpforms)

    target_filename = (
        f"./out/{year} - JPOS Registration Report {week_start} to {week_end}.xlsx"
    )

    report.write_excel(
        workbook=target_filename,
        worksheet="Registrations {year}-W{week}",
        table_style="Table Style Medium 2",
    )

    print(f"Report generated! Location: {target_filename}")


if __name__ == "__main__":
    generate_report(WOOCOMMERCE_PATH, WPFORMS_PATH, YEAR, WEEK)
