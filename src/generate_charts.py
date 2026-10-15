from pathlib import Path
import matplotlib.pyplot as plt

from analyze_data import (
    load_data,
    calculate_analysis,
)


BASE_DIR = Path(__file__).resolve().parent.parent
CHART_DIR = BASE_DIR / "charts"


COLORS = {
    "North America": "#1F4E78",
    "Europe": "#5B9BD5",
    "APAC": "#70AD47",
    "Latin America": "#C00000",
}


def style_axis(axis, title, y_label):

    axis.set_title(
        title,
        fontsize=15,
        fontweight="bold",
        pad=15,
    )

    axis.set_ylabel(
        y_label,
        fontsize=10,
    )

    axis.grid(
        axis="y",
        alpha=0.25,
    )

    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)


def create_revenue_chart(performance):

    chart_data = performance.pivot(
        index="Quarter",
        columns="Region",
        values="Revenue_USD_M",
    )

    chart_data = chart_data.reindex(
        [
            "2025-Q3",
            "2025-Q4",
            "2026-Q1",
            "2026-Q2",
        ]
    )

    figure, axis = plt.subplots(
        figsize=(10, 5.2)
    )

    for region in chart_data.columns:

        axis.plot(
            chart_data.index,
            chart_data[region],
            marker="o",
            linewidth=2.5,
            markersize=6,
            label=region,
            color=COLORS[region],
        )

    style_axis(
        axis,
        "Quarterly Revenue by Region",
        "Revenue (USD millions)",
    )

    axis.legend(
        frameon=False,
        ncol=2,
        loc="upper left",
    )

    figure.tight_layout()

    output_path = (
        CHART_DIR / "revenue_by_region.png"
    )

    figure.savefig(
        output_path,
        dpi=200,
        bbox_inches="tight",
        facecolor="white",
    )

    plt.close(figure)

    return output_path


def create_nps_chart(customer):

    chart_data = customer.pivot(
        index="Quarter",
        columns="Region",
        values="NPS",
    )

    chart_data = chart_data.reindex(
        [
            "2025-Q3",
            "2025-Q4",
            "2026-Q1",
            "2026-Q2",
        ]
    )

    figure, axis = plt.subplots(
        figsize=(10, 5.2)
    )

    for region in chart_data.columns:

        axis.plot(
            chart_data.index,
            chart_data[region],
            marker="o",
            linewidth=2.5,
            markersize=6,
            label=region,
            color=COLORS[region],
        )

    style_axis(
        axis,
        "NPS Trend by Region",
        "Net Promoter Score",
    )

    axis.legend(
        frameon=False,
        ncol=2,
        loc="upper left",
    )

    figure.tight_layout()

    output_path = (
        CHART_DIR / "nps_by_region.png"
    )

    figure.savefig(
        output_path,
        dpi=200,
        bbox_inches="tight",
        facecolor="white",
    )

    plt.close(figure)

    return output_path


def create_risk_chart(risk):

    latest = risk[
        risk["Quarter"].astype(str)
        == "2026-Q2"
    ].copy()

    latest = latest.sort_values(
        "Stockout_Rate_pct",
        ascending=True,
    )

    colors = [
        COLORS[region]
        for region in latest["Region"]
    ]

    figure, axis = plt.subplots(
        figsize=(10, 5.2)
    )

    bars = axis.barh(
        latest["Region"],
        latest["Stockout_Rate_pct"],
        color=colors,
    )

    axis.set_title(
        "Q2 2026 Stockout Rate by Region",
        fontsize=15,
        fontweight="bold",
        pad=15,
    )

    axis.set_xlabel(
        "Stockout Rate (%)",
        fontsize=10,
    )

    axis.grid(
        axis="x",
        alpha=0.25,
    )

    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)

    for bar, value in zip(
        bars,
        latest["Stockout_Rate_pct"],
    ):

        axis.text(
            value + 0.2,
            bar.get_y()
            + bar.get_height() / 2,
            f"{value:.1f}%",
            va="center",
            fontsize=10,
            fontweight="bold",
        )

    figure.tight_layout()

    output_path = (
        CHART_DIR / "stockout_rate.png"
    )

    figure.savefig(
        output_path,
        dpi=200,
        bbox_inches="tight",
        facecolor="white",
    )

    plt.close(figure)

    return output_path


def main():

    print("=" * 60)
    print("GENERATING NOVARETAIL CHARTS")
    print("=" * 60)

    CHART_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    data = load_data()

    (
        results,
        performance,
        customer,
        risk,
    ) = calculate_analysis(data)

    revenue_chart = create_revenue_chart(
        performance
    )

    nps_chart = create_nps_chart(
        customer
    )

    risk_chart = create_risk_chart(
        risk
    )

    print("\nCharts created successfully:")

    print(f"1. {revenue_chart}")
    print(f"2. {nps_chart}")
    print(f"3. {risk_chart}")


if __name__ == "__main__":
    main()