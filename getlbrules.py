# RETAILTECH TAREGT GROUP ARN REPLACE SCRIPT

import boto3
import openpyxl

# AWS credentials and Load Balancer ARN
LOAD_BALANCER_ARN = "arn:aws:elasticloadbalancing:ap-south-1:765020061369:loadbalancer/app/gopay-orders-uat-lb/63c074c021fd1e78"
REGION_NAME = "ap-south-1"

# Initialize AWS client
client = boto3.client('elbv2', region_name=REGION_NAME)

# Initialize target group client
tg_client = boto3.client('elbv2', region_name=REGION_NAME)

# Load the Excel workbook
wb = openpyxl.load_workbook('loadbalancer_rules.xlsx')
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

# Skip header row
for row in sheet.iter_rows(min_row=2, values_only=True):
    listener_port = row[0]
    priority = row[1]  # If auto-assigning, this will be ignored
    action_type = row[3]
    target_group_name = row[4]  # Now reading the target group name instead of ARN
    condition_field = row[5]
    condition_value = row[6]

    print(f"Creating rule for port {listener_port}...")

    # Find the listener ARN for the given port
    listeners = client.describe_listeners(LoadBalancerArn=LOAD_BALANCER_ARN)['Listeners']
    listener_arn = next((l['ListenerArn'] for l in listeners if l['Port'] == listener_port), None)

    if not listener_arn:
        print(f"No listener found on port {listener_port}. Skipping...")
        continue

    # Automatically get the next available priority
    priority = get_next_available_priority(listener_arn)
    print(f"Assigned priority {priority}.")

    # Get the target group ARN from the target group name
    try:
        target_group_arn = get_target_group_arn(target_group_name)
    except ValueError as e:
        print(e)
        continue

    # Create the conditions
    conditions = [{
        'Field': condition_field,
        'Values': [condition_value]
    }]

    # Create the actions
    if action_type == "redirect":
        # Add RedirectConfig for redirect actions
        actions = [{
            'Type': 'redirect',
            'RedirectConfig': {
                'Protocol': 'HTTPS',  # You can adjust the protocol, port, path etc.
                'Port': '443',
                'StatusCode': 'HTTP_301'
            }
        }]
    else:
        actions = [{
            'Type': action_type,
            'TargetGroupArn': target_group_arn
        }]

    # Create the rule
    try:
        response = client.create_rule(
            ListenerArn=listener_arn,
            Priority=int(priority),  # Automatically assigned
            Conditions=conditions,
            Actions=actions
        )
        new_rule_arn = response['Rules'][0]['RuleArn']
        print(f"Rule created successfully with ARN: {new_rule_arn} for priority {priority} on listener {listener_arn}.")
    except Exception as e:
        print(f"Failed to create rule for priority {priority}: {e}")

print("Finished creating rules.")

