import boto3
import openpyxl

# AWS Load Balancer ARN
LOAD_BALANCER_ARN = "arn:aws:elasticloadbalancing:ap-south-1:092042625037:loadbalancer/app/central-platforms-nonprod-lb/302284ca38d805dc"

# Create a boto3 client for ELBv2
client = boto3.client('elbv2')

# Create a new Excel workbook
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Load Balancer Rules"

# Define the headers for the Excel sheet
headers = ["Listener Port", "Rule Priority", "Rule ARN", "Action Type", "Target Group ARN", "Redirect URL", "Condition Field", "Condition Value"]
sheet.append(headers)

# Fetch the listeners for the specified load balancer
listeners = client.describe_listeners(LoadBalancerArn=LOAD_BALANCER_ARN)['Listeners']

# Iterate over each listener
for listener in listeners:
    listener_arn = listener['ListenerArn']
    port = listener['Port']
    
    # Only process listeners on port 80 or 443
    if port in [80, 443]:
        print(f"Fetching rules for listener on port: {port}")

        # Fetch the rules for each listener
        rules = client.describe_rules(ListenerArn=listener_arn)['Rules']

        # Iterate over each rule
        for rule in rules:
            rule_arn = rule['RuleArn']
            priority = rule['Priority']
            actions = rule['Actions']
            conditions = rule['Conditions']

            # Iterate over actions
            for action in actions:
                action_type = action['Type']
                target_group_arn = action.get('TargetGroupArn', "N/A")
                redirect_url = "N/A"

                # Check if action is a redirect and extract URL if applicable
                if action_type == 'redirect':
                    redirect_url = action.get('RedirectConfig', {}).get('Url', "N/A")

                # Iterate over conditions
                for condition in conditions:
                    field = condition.get('Field', "N/A")
                    values = condition.get('Values', ["N/A"])

                    # If there are multiple values, save each one in a new row
                    for value in values:
                        row = [port, priority, rule_arn, action_type, target_group_arn, redirect_url, field, value]
                        sheet.append(row)

# Auto-adjust column widths for better readability
for col in sheet.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column name
    for cell in col:
        try:  # Necessary to avoid issues with numbers and NoneType cells
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    sheet.column_dimensions[column].width = adjusted_width

# Save the workbook
output_file = "loadbalancer_rules.xlsx"
wb.save(output_file)

print(f"Rules saved to {output_file}")
