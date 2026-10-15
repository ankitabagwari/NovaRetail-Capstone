
The process works as follows:

Template Inspection
inspect_template.py opens the provided PowerPoint template and identifies the existing shapes on all five slides using their intentional shape names. This ensures that the original slide structure and layout are preserved.

Data Loading and Validation
analyze_data.py reads the NovaRetail Excel workbook using pandas and openpyxl. It validates the required sheets and columns before performing any calculations.

Data Analysis and KPI Calculation
The program analyzes Regional Performance, Customer Experience, and Operational Risk data. It calculates the executive KPIs required for the presentation, including total revenue, quarter-over-quarter revenue growth, weighted customer NPS, and total units sold.

Chart Generation
generate_charts.py uses matplotlib to create charts directly from the dataset. The generated charts include:

Revenue trend by region
NPS trend by region
Q2 2026 stockout rate by region
Business Story and Insights
The analyzed data is used to create insight-led headlines, executive-summary content, performance observations, customer-experience findings, operational-risk findings, and recommended actions.

PowerPoint Generation
generate_presentation.py opens the original PowerPoint template and locates the required areas using the existing shape names such as title_2, chart_performance, insights_customer, and chart_risk.

Charts and Icons
The generated charts are inserted into their corresponding chart placeholders using the placeholder's original position and dimensions. Relevant icons from the provided icon repository are also inserted according to each slide's business theme.

Template Preservation
The program keeps the original five-slide structure and uses the existing template layout. The supplied template remains unchanged because the completed presentation is saved as a separate file.

Final Output
The completed presentation is saved as:

output/NovaRetail_QBR_Final.pptx

End-to-End Execution
main.py brings the individual steps together so the complete workflow can be executed from the project root using:

python src/main.py

Running the project again regenerates the analysis, charts, and final PowerPoint from the supplied source files.




To regenerate the presentation:

-Place the NovaRetail Excel dataset in the assets directory.
-Keep the original PowerPoint template in the assets directory.
-Keep the icon image files under assets/icons.
- Create all folder/subfolder/files as created
-Install the required packages using requirements.txt.
- Keep the code as it is in inspect_template
-Run: analyze_data , generate_charts, generate_presentation or directly run main.py
- Output will come under  NovaRetail_QBR_Final.pptx
 