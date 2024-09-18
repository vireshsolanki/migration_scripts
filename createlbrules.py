import boto3
import openpyxl
import time

# AWS credentials and Load Balancer ARN
LOAD_BALANCER_ARN = "arn:aws:elasticloadbalancing:ap-south-1:145023133364:loadbalancer/app/central-platforms-nonprod-lb/f1c2f465fff37a80"
REGION_NAME = "ap-south-1"

# Initialize AWS client
client = boto3.client('elbv2', region_name=REGION_NAME)
tg_client = boto3.client('elbv2', region_name=REGION_NAME)

# Load the Excel workbook
wb = openpyxl.load_workbook("loadbalancer_rules.xlsx")
sheet = wb.active

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
        # Sort rules by priority to delete the oldest first
        sorted_rules = sorted(existing_rules, key=lambda r: int(r['Priority']))
        for rule in sorted_rules:
            if len(existing_rules) >= max_rules:
                client.delete_rule(RuleArn=rule['RuleArn'])
                print(f"Deleted rule {rule['RuleArn']}")
                # Re-fetch the rules list after deletion
                existing_rules = client.describe_rules(ListenerArn=listener_arn)['Rules']
            else:
                break

valid_condition_fields = ['http-header', 'http-request-method', 'host-header', 'query-string', 'source-ip', 'path-pattern']

# Skip header row
for row in sheet.iter_rows(min_row=2, values_only=True):
    listener_port = row[0]
    action_type = row[3]
    target_group_name = row[4]  # Now reading the target group name instead of ARN
    condition_field = row[6]
    condition_value = row[7]

    # Skip invalid condition fields
    if condition_field == 'N/A' or condition_field not in valid_condition_fields:
        print(f"Skipping rule creation for port {listener_port} due to invalid condition field '{condition_field}'.")
        continue

    print(f"Creating rule for port {listener_port}...")

    # Find the listener ARN for the given port
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

    # Get the target group ARN from the target group name
    if action_type != "redirect":
        try:
            target_group_arn = get_target_group_arn(target_group_name)
        except ValueError as e:
            print(e)
            continue
        actions = [{
            'Type': action_type,
            'TargetGroupArn': target_group_arn
        }]
    else:
        actions = [{
            'Type': 'redirect',
            'RedirectConfig': {
                'Protocol': 'HTTPS',
                'Port': '443',
                'StatusCode': 'HTTP_301'
            }
        }]

    # Create the conditions
    conditions = [{
        'Field': condition_field,
        'Values': [condition_value]
    }]

    # Create the rule
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
