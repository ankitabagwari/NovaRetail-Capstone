from inspect_template import inspect_template
from analyze_data import main as run_analysis
from generate_charts import main as run_charts
from generate_presentation import generate_presentation


def main():
    print("=" * 70)
    print("NOVARETAIL CAPSTONE - EXECUTIVE DECK AUTOMATION")
    print("=" * 70)

    # Step 1: Inspect the PowerPoint template
    print("\nSTEP 1: Inspecting PowerPoint template...")
    inspect_template()

    # Step 2: Analyze the Excel dataset
    print("\nSTEP 2: Analyzing NovaRetail dataset...")
    run_analysis()

    # Step 3: Generate charts
    print("\nSTEP 3: Generating charts...")
    run_charts()

    # Step 4: Generate final PowerPoint
    print("\nSTEP 4: Generating final PowerPoint...")
    generate_presentation()

    print("\n" + "=" * 70)
    print("NOVARETAIL CAPSTONE COMPLETED SUCCESSFULLY")
    print("=" * 70)

    print(
        "\nFinal presentation created at:"
        "\noutput/NovaRetail_QBR_Final.pptx"
    )


if __name__ == "__main__":
    main()