
## **Load Balancer Rules Management Guide**

This guide walks you through the process of generating an Excel sheet containing load balancer rules from the source environment and applying those rules to the destination environment using two scripts: `Getlb.py` and `Createlb.py`.

### **Requirements**

Before running the scripts, ensure you have the following installed:

- Python 3.x
- AWS CLI configured with appropriate credentials for both the source and destination accounts
- Required Python packages:
  ```bash
  pip3 install boto3 openpyxl
  ```

---

### **Step 1: Generate the Excel Sheet in the Source Environment**

**Objective**: Export load balancer rules from the **source** environment.

1. **Run the Script `Getlb.py`**:
   - The script retrieves load balancer information from the AWS source account and generates an Excel sheet.

   **Instructions**:
   1. Install the necessary dependencies:
      ```bash
      pip3 install boto3 openpyxl
      ```
   2. Execute the script in the **source** environment:
      ```bash
      python3 Getlb.py
      ```
   3. After running the script, it will create an Excel file named `loadbalancer_rules.xlsx` in the working directory. This file contains the relevant information for creating load balancer rules.

2. **Modify Domain Names**:
   - If required, open the `loadbalancer_rules.xlsx` file and modify the domain names or other necessary fields as per your destination environment requirements.

   **Instructions**:
   1. Open the file in Excel or any compatible spreadsheet application.
   2. Locate the columns containing domain names or other fields that need to be updated.
   3. Make the necessary changes and save the file.

---

### **Step 2: Apply Load Balancer Rules in the Destination Environment**

**Objective**: Import the updated rules into the **destination** environment by creating load balancer rules.

1. **Run the Script `Createlb.py`**:
   - This script reads the Excel sheet and creates the appropriate load balancer rules in the destination AWS account.

   **Instructions**:
   1. Install the necessary dependencies in the **destination** environment:
      ```bash
      pip3 install boto3 openpyxl
      ```
   2. **Update the `Createlb.py` script**:
      - Before running the script, ensure the following:
        - **Load Balancer ARN**: Update the `LOAD_BALANCER_ARN` variable in the script to match the ARN of the load balancer in the destination environment.
        - **Target Groups**: Verify that the target group names in the Excel file are valid and exist in the destination AWS account.
   
   3. Execute the script in the **destination** environment:
      ```bash
      python3 Createlb.py
      ```

2. The script will process the `loadbalancer_rules.xlsx` file and create the rules in the destination load balancer accordingly. Make sure that your AWS credentials in the destination environment have the necessary permissions to perform these operations (e.g., creating rules, modifying load balancers).

---

### **Script Details**

- **`Getlb.py`**:
  - **Purpose**: Retrieves information about load balancers from the source environment and generates an Excel sheet with the details.
  - **Output**: The Excel file `loadbalancer_rules.xlsx`.

- **`Createlb.py`**:
  - **Purpose**: Reads the Excel sheet and creates load balancer rules in the destination environment based on the provided data.
  - **Configuration**: Before running, update the `LOAD_BALANCER_ARN` and ensure that the target group names match those in the destination AWS account.

---

### **Important Notes**:

- Always make sure that the AWS credentials used in both the source and destination environments have sufficient permissions to list and create load balancer rules, respectively.
- It's a good practice to take backups or snapshots of your existing load balancer configurations before applying new rules, especially in production environments.

---
