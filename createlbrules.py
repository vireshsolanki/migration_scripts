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
file_path = "/mnt/data/loadbalancer_rules_combined.xlsx"
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

# List of valid condition fields
valid_condition_fields = ['http-header', 'http-request-method', 'host-header', 'query-string', 'source-ip', 'path-pattern']

# Get existing target groups
target_groups = tg_client.describe_target_groups()['TargetGroups']
target_groups_dict = {tg['TargetGroupName']: tg['TargetGroupArn'] for tg in target_groups}

def get_next_available_priority(listener_arn):
    """Fetch existing rules and return the next available priority."""
    existing_rules = client.describe_rules(ListenerArn=listener_arn)['Rules']
    used_priorities = [int(rule['Priority']) for rule in existing_rules if rule['Priority'].isdigit()]
    return max(used_priorities) + 1 if used_priorities else 1

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

# Dictionary to track combined conditions for each listener port
listener_conditions = {}

# Skip header row and process each rule from the Excel file
for row in sheet.iter_rows(min_row=2, values_only=True):
    listener_port = row[0]  # Listener port (80 or 443)
    action_type = row[3]  # Action type (forward or redirect)
    target_group_name = row[5]  # Target group name
    condition_field = row[7]  # Condition field (host-header, etc.)
    condition_value = row[8]  # Condition value (e.g., domain name)

    # Skip invalid or N/A condition fields
    if condition_field == 'N/A' or condition_field not in valid_condition_fields:
        print(f"Skipping rule creation for port {listener_port} due to invalid condition field '{condition_field}'.")
        continue

    # Initialize the condition combination for the listener port if not present
    if listener_port not in listener_conditions:
        listener_conditions[listener_port] = {'host-header': [], 'path-pattern': []}

    # Add conditions to the respective fields
    if condition_field in ['host-header', 'path-pattern']:
        listener_conditions[listener_port][condition_field].append(condition_value)

    # Check if we have both host-header and path-pattern conditions ready to create a combined rule
    if listener_conditions[listener_port]['host-header'] and listener_conditions[listener_port]['path-pattern']:
        print(f"Creating combined rule for port {listener_port}...")

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
            target_group_arn = target_groups_dict.get(target_group_name, None)
            if not target_group_arn:
                print(f"Target group with name {target_group_name} not found. Skipping...")
                continue
            actions = [{'Type': action_type, 'TargetGroupArn': target_group_arn}]
        else:
            actions = [{'Type': 'redirect', 'RedirectConfig': {'Protocol': 'HTTPS', 'Port': '443', 'StatusCode': 'HTTP_301'}}]

        # Combine host-header and path-pattern into a single condition
        conditions = [
            {'Field': 'host-header', 'Values': listener_conditions[listener_port]['host-header']},
            {'Field': 'path-pattern', 'Values': listener_conditions[listener_port]['path-pattern']}
        ]

        # Create the combined rule
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

        # Clear the conditions after creating the combined rule
        listener_conditions[listener_port] = {'host-header': [], 'path-pattern': []}

        # Add a delay before creating the next rule
        time.sleep(2)  # Delay in seconds

print("Finished creating rules.")
