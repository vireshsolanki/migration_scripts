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
headers = ["Listener Port", "Rule Priority", "Rule ARN", "Action Type", "Target Group ARN", "Target Group Name", "Redirect URL", "Condition Field", "Condition Value"]
sheet.append(headers)

# Fetch the listeners for the specified load balancer
listeners = client.describe_listeners(LoadBalancerArn=LOAD_BALANCER_ARN)['Listeners']

# Fetch the target groups
target_groups = client.describe_target_groups()['TargetGroups']
target_groups_dict = {tg['TargetGroupArn']: tg['TargetGroupName'] for tg in target_groups}

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
            
            # Dictionary to combine conditions by target group ARN
            condition_combinations = {}

            # Iterate over actions
            for action in actions:
                action_type = action['Type']
                target_group_arn = action.get('TargetGroupArn', "N/A")
                target_group_name = target_groups_dict.get(target_group_arn, "N/A")
                redirect_url = "N/A"

                # Check if action is a redirect and extract URL if applicable
                if action_type == 'redirect':
                    redirect_url = action.get('RedirectConfig', {}).get('Url', "N/A")

                # Initialize condition_combinations for this target group if not present
                if target_group_arn not in condition_combinations:
                    condition_combinations[target_group_arn] = {
                        'host-header': [],
                        'path-pattern': [],
                        'others': []
                    }

                # Iterate over conditions
                for condition in conditions:
                    field = condition.get('Field', "N/A")
                    values = condition.get('Values', ["N/A"])

                    # Store values under appropriate condition type
                    if field in ['host-header', 'path-pattern']:
                        condition_combinations[target_group_arn][field].extend(values)
                    else:
                        # Store other types of conditions
                        for value in values:
                            condition_combinations[target_group_arn]['others'].append({'Field': field, 'Value': value})

            # Combine conditions and add rows to the Excel sheet
            for target_group_arn, combined_conditions in condition_combinations.items():
                # Combine host-header and path-pattern conditions if both are present
                if combined_conditions['host-header'] and combined_conditions['path-pattern']:
                    combined_value = (
                        f"Host Headers: {', '.join(combined_conditions['host-header'])} | "
                        f"Path Patterns: {', '.join(combined_conditions['path-pattern'])}"
                    )
                    row = [port, priority, rule_arn, action_type, target_group_arn, target_group_name, redirect_url, 'host-header + path-pattern', combined_value]
                    sheet.append(row)
                else:
                    # Add host-header conditions if present
                    if combined_conditions['host-header']:
                        for value in combined_conditions['host-header']:
                            row = [port, priority, rule_arn, action_type, target_group_arn, target_group_name, redirect_url, 'host-header', value]
                            sheet.append(row)

                    # Add path-pattern conditions if present
                    if combined_conditions['path-pattern']:
                        for value in combined_conditions['path-pattern']:
                            row = [port, priority, rule_arn, action_type, target_group_arn, target_group_name, redirect_url, 'path-pattern', value]
                            sheet.append(row)

                    # Add other conditions if present
                    for condition in combined_conditions['others']:
                        row = [port, priority, rule_arn, action_type, target_group_arn, target_group_name, redirect_url, condition['Field'], condition['Value']]
                        sheet.append(row)

# Auto-adjust column widths for better readability
for col in sheet.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column name
    for cell in col:
        try:  # Necessary to avoid issues with numbers and NoneType cells
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2)
    sheet.column_dimensions[column].width = adjusted_width

# Save the workbook
output_file = "loadbalancer_rules_combined.xlsx"
wb.save(output_file)

print(f"Rules saved to {output_file}")
