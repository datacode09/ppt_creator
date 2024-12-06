from pptx.enum.shapes import MSO_CONNECTOR

# Corrected method to add arrows
def add_arrow(start_shape, end_shape):
    # Calculate start and end points for the arrow
    start_x = start_shape.left + start_shape.width // 2
    start_y = start_shape.top + start_shape.height
    end_x = end_shape.left + end_shape.width // 2
    end_y = end_shape.top
    # Add a connector (arrow)
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, start_x, start_y, end_x, end_y)
    # Format the connector as an arrow
    connector.line.end_arrowhead = True

# Reconnect the steps
add_arrow(shapes["Start"], shapes["Initialize Macro Variables"])
add_arrow(shapes["Initialize Macro Variables"], shapes["Include Setup File"])
add_arrow(shapes["Include Setup File"], shapes["Define Date Range"])
add_arrow(shapes["Define Date Range"], shapes["Run Macro: passthrough_date"])
add_arrow(shapes["Run Macro: passthrough_date"], shapes["Iterate Through Dates"])
add_arrow(shapes["Iterate Through Dates"], shapes["Include Entity Creation File"])
add_arrow(shapes["Include Entity Creation File"], shapes["Store Final Entity Data"])
add_arrow(shapes["Store Final Entity Data"], shapes["Update Date (Next Month)"])
add_arrow(shapes["Update Date (Next Month)"], shapes["End Iteration"])
add_arrow(shapes["End Iteration"], shapes["End"])

# Save the presentation
output_path = "/mnt/data/SAS_Code_Workflow_Flowchart_Corrected.pptx"
prs.save(output_path)

output_path
