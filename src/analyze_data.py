from pathlib import Path
import json
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent.parent

DATASET_PATH = (
    BASE_DIR
    / "assets"
    / "novaretail_dataset.xlsx"
)

ANALYSIS_PATH = (
    BASE_DIR
    / "output"
    / "analysis_results.json"
)


SHEET_OPTIONS = {
    "performance": [
        "Regional_Performance",
        "Regional Performance",
    ],
    "customer": [
        "Customer_Experience",
        "Customer Experience",
        "Customer Exp",
    ],
    "risk": [
        "Operational_Risk",
        "Operational Risk",
    ],
}


REQUIRED_COLUMNS = {
    "performance": [
        "Quarter",
        "Region",
        "Revenue_USD_M",
        "Units_Sold_K",
        "Returns_Rate_pct",
        "Marketing_Spend_USD_M",
    ],
    "customer": [
        "Quarter",
        "Region",
        "NPS",
        "CSAT_pct",
        "Avg_Delivery_Days",
        "Support_Tickets_K",
    ],
    "risk": [
        "Quarter",
        "Region",
        "Stockout_Rate_pct",
        "Return_Processing_Days",
        "Warehouse_Capacity_Util_pct",
        "Fraud_Incidents",
    ],
}


QUARTER_ORDER = [
    "2025-Q3",
    "2025-Q4",
    "2026-Q1",
    "2026-Q2",
]


def find_sheet(workbook, possible_names):
    for name in possible_names:
        if name in workbook:
            return workbook[name].copy()

    raise ValueError(
        f"None of these sheets were found: {possible_names}"
    )


def load_data():
    """Load and validate all three NovaRetail sheets."""

    if not DATASET_PATH.exists():
        raise FileNotFoundError(
            f"Dataset not found: {DATASET_PATH}"
        )

    workbook = pd.read_excel(
        DATASET_PATH,
        sheet_name=None,
        engine="openpyxl",
    )

    print("\nSheets found:")

    for name in workbook:
        print(f"  - {name}")

    data = {}

    for logical_name, possible_names in SHEET_OPTIONS.items():

        dataframe = find_sheet(
            workbook,
            possible_names,
        )

        dataframe.columns = [
            str(column).strip()
            for column in dataframe.columns
        ]

        missing_columns = (
            set(REQUIRED_COLUMNS[logical_name])
            - set(dataframe.columns)
        )

        if missing_columns:
            raise ValueError(
                f"{logical_name} is missing columns: "
                f"{sorted(missing_columns)}"
            )

        data[logical_name] = dataframe

        print(
            f"Loaded {logical_name}: "
            f"{len(dataframe)} rows"
        )

    return data


def prepare_quarters(dataframe):
    """Put quarters into chronological order."""

    dataframe = dataframe.copy()

    dataframe["Quarter"] = (
        dataframe["Quarter"]
        .astype(str)
        .str.strip()
    )

    dataframe["Quarter"] = pd.Categorical(
        dataframe["Quarter"],
        categories=QUARTER_ORDER,
        ordered=True,
    )

    dataframe = dataframe.sort_values(
        ["Quarter", "Region"]
    )

    return dataframe


def calculate_analysis(data):
    """Calculate executive KPIs used in the presentation."""

    performance = prepare_quarters(
        data["performance"]
    )

    customer = prepare_quarters(
        data["customer"]
    )

    risk = prepare_quarters(
        data["risk"]
    )

    latest_quarter = "2026-Q2"
    previous_quarter = "2026-Q1"

    current_performance = performance[
        performance["Quarter"].astype(str)
        == latest_quarter
    ].copy()

    previous_performance = performance[
        performance["Quarter"].astype(str)
        == previous_quarter
    ].copy()

    current_customer = customer[
        customer["Quarter"].astype(str)
        == latest_quarter
    ].copy()

    total_revenue = float(
        current_performance[
            "Revenue_USD_M"
        ].sum()
    )

    previous_revenue = float(
        previous_performance[
            "Revenue_USD_M"
        ].sum()
    )

    qoq_growth = (
        (total_revenue - previous_revenue)
        / previous_revenue
        * 100
    )

    total_units = float(
        performance[
            "Units_Sold_K"
        ].sum()
    )

    customer_with_units = (
        current_customer.merge(
            current_performance[
                [
                    "Region",
                    "Units_Sold_K",
                ]
            ],
            on="Region",
            how="left",
        )
    )

    weighted_nps = (
        (
            customer_with_units["NPS"]
            * customer_with_units[
                "Units_Sold_K"
            ]
        ).sum()
        / customer_with_units[
            "Units_Sold_K"
        ].sum()
    )

    results = {
        "latest_quarter": latest_quarter,

        "total_revenue_usd_m": round(
            total_revenue,
            1,
        ),

        "previous_revenue_usd_m": round(
            previous_revenue,
            1,
        ),

        "qoq_growth_pct": round(
            qoq_growth,
            1,
        ),

        "weighted_nps": round(
            float(weighted_nps),
            1,
        ),

        "total_units_sold_k_all_quarters": round(
            total_units,
            1,
        ),
    }

    return (
        results,
        performance,
        customer,
        risk,
    )


def save_analysis(results):

    ANALYSIS_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    ANALYSIS_PATH.write_text(
        json.dumps(
            results,
            indent=4,
        ),
        encoding="utf-8",
    )


def main():

    print("=" * 60)
    print("NOVARETAIL DATA ANALYSIS")
    print("=" * 60)

    data = load_data()

    (
        results,
        _,
        _,
        _,
    ) = calculate_analysis(data)

    save_analysis(results)

    print("\nAnalysis results:")

    for key, value in results.items():
        print(f"{key}: {value}")

    print(
        "\nAnalysis completed successfully."
    )

    print(
        f"Saved to: {ANALYSIS_PATH}"
    )


if __name__ == "__main__":
    main()