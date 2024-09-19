import boto3
import openpyxl
import time

# AWS credentials and Load Balancer ARN
LOAD_BALANCER_ARN = "arn:aws:elasticloadbalancing:ap-south-1:092042625037:loadbalancer/app/cloud-dev-ondemand-lb/04f65db33b0e9ca8"
REGION_NAME = "ap-south-1"

# Initialize AWS client
client = boto3.client('elbv2', region_name=REGION_NAME)
tg_client = boto3.client('elbv2', region_name=REGION_NAME)

# Load the Excel workbook
wb = openpyxl.load_workbook("loadbalancer_rules.xlsx")
sheet = wb.active

# List of valid condition fields
valid_condition_fields = ['http-header', 'http-request-method', 'host-header', 'query-string', 'source-ip', 'path-pattern']

def get_next_available_priority(listener_arn):
    """Fetch existing rules and return the next available priority."""
    existing_rules = client.describe_rules(ListenerArn=listener_arn)['Rules']
    used_priorities = [int(rule['Priority']) for rule in existing_rules if rule['Priority'].isdigit()]
    return max(used_priorities) + 1 if used_priorities else 1

def get_target_group_arn(target_group_name):
    """Get the ARN of a target group by its name."""
    target_groups = tg_client.describe_target_groups()['TargetGroups']
    target_group_arn = next((tg['TargetGroupArn'] for tg in target_groups if tg['TargetGroupName'] == target_group_name), None)
    if not target_group_arn:
        raise ValueError(f"Target group with name {target_group_name} not found.")
    return target_group_arn

def clean_up_rules(listener_arn, max_rules=1000):
    """Ensure the number of rules does not exceed the maximum limit."""
    existing_rules = client.describe_rules(ListenerArn=listener_arn)['Rules']
    if len(existing_rules) >= max_rules:
        print(f"Maximum number of rules reached on listener {listener_arn}. Removing old rules...")
        sorted_rules = sorted(existing_rules, key=lambda r: int(r['Priority']))
        for rule in sorted_rules:
            if len(existing_rules) >= max_rules:
                client.delete_rule(RuleArn=rule['RuleArn'])
                print(f"Deleted rule {rule['RuleArn']}")
                existing_rules = client.describe_rules(ListenerArn=listener_arn)['Rules']
            else:
                break

# Function to group conditions based on listener port and action type
def group_conditions(sheet):
    grouped_conditions = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        listener_port = row[0]
        action_type = row[3]
        target_group_name = row[5]
        condition_field = row[7]
        condition_value = row[8]

        if condition_field == 'N/A' or condition_field not in valid_condition_fields:
            continue

        key = (listener_port, action_type, target_group_name)
        if key not in grouped_conditions:
            grouped_conditions[key] = []

        grouped_conditions[key].append({'Field': condition_field, 'Values': [condition_value]})
    
    return grouped_conditions

# Get grouped conditions
grouped_conditions = group_conditions(sheet)

# Iterate through grouped conditions to create rules
for key, conditions in grouped_conditions.items():
    listener_port, action_type, target_group_name = key

    print(f"Creating rule for port {listener_port} with combined conditions...")

    # Find the listener ARN for the given port (80 or 443)
    listeners = client.describe_listeners(LoadBalancerArn=LOAD_BALANCER_ARN)['Listeners']
    listener_arn = next((l['ListenerArn'] for l in listeners if l['Port'] == listener_port), None)

    if not listener_arn:
        print(f"No listener found on port {listener_port}. Skipping...")
        continue

    # Clean up old rules if necessary
    clean_up_rules(listener_arn)

    # Automatically get the next available priority
    priority = get_next_available_priority(listener_arn)
    print(f"Assigned priority {priority}.")

    # Get the target group ARN from the target group name (if not a redirect rule)
    if action_type != "redirect":
        try:
            target_group_arn = get_target_group_arn(target_group_name)
        except ValueError as e:
            print(e)
            continue
        actions = [{'Type': action_type, 'TargetGroupArn': target_group_arn}]
    else:
        actions = [{'Type': 'redirect', 'RedirectConfig': {'Protocol': 'HTTPS', 'Port': '443', 'StatusCode': 'HTTP_301'}}]

    # Create the rule with combined conditions
    try:
        response = client.create_rule(
            ListenerArn=listener_arn,
            Priority=int(priority),
            Conditions=conditions,
            Actions=actions
        )
        new_rule_arn = response['Rules'][0]['RuleArn']
        print(f"Rule created successfully with ARN: {new_rule_arn} for priority {priority} on listener {listener_arn}.")
    except Exception as e:
        print(f"Failed to create rule for priority {priority}: {e}")

    # Add a delay before creating the next rule
    time.sleep(2)  # Delay in seconds

print("Finished creating rules.")
