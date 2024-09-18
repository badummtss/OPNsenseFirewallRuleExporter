# OPNsense Firewall Rule Exporter

This repository contains a Python script that extracts OPNsense firewall rules from an XML configuration file and writes the rules into an Excel file. The Excel file organizes the firewall rules into separate sheets based on the interface, with additional functionality to handle aliases and comments.

## Features

- **Interface-based Sheet Organization**: The script creates a separate sheet in the Excel file for each network interface found in the OPNsense firewall configuration.
- **Alias Detection and Comments**: If aliases are used in source, destination, or port fields, the script automatically adds these aliases as comments in the Excel file.
- **Rule Fields**: For each rule, the following fields are extracted and displayed:
  - `Aktiviert?`: Whether the rule is enabled or disabled.
  - `Action`: Whether the rule allows (`pass`) or blocks (`block`) traffic.
  - `Protocol`: The protocol used in the rule (e.g., TCP, UDP, any).
  - `Source` and `Port`: The source network or alias and the port.
  - `Destination` and `Port`: The destination network or alias and the port.
  - `Gateway`: The gateway associated with the rule, if any.
  - `Schedule`: The schedule associated with the rule, if any.
  - `Number of Rule`: The rule number or tracker.
  - `Description`: The description of the rule.
  - `Rules`: Any additional rules or custom rules for each interface.

## Prerequisites

- **Python 3.7+**: Ensure you have Python installed on your machine.
- **Libraries**:
  - `openpyxl`: Used to generate and manipulate Excel files.
  - `xml.etree.ElementTree`: Used for parsing the XML file.
  
  You can install the required libraries using pip:

  ```bash
  pip install openpyxl
## Usage

1. **Clone the repository**:

   ```bash
   git clone https://github.com/DominikWoh/opnsense-firewall-exporter.git
   cd opnsense-firewall-exporter
Place the OPNsense XML Configuration: Copy the OPNsense configuration XML file into the repository folder.

### Configuration

Update the following in the script:

1. **XML File Path**:
   ```python
   xml_file = r'C:\GitHub\opnsense\config-opnsense.xml'
Replace `'C:\GitHub\opnsense\config-opnsense.xml'` with the actual path to your OPNsense XML file.

2. **Excel Output Path**:

    ```python
    output_file = 'firewall_rules_by_interface.xlsx'
Replace 'firewall_rules_by_interface.xlsx' with your desired output file name and location.

### Run the script:

    python opnsense_exporter.py

### Generated Excel File

After running the script, an Excel file named `firewall_rules_by_interface.xlsx` will be generated in the current directory. This file will contain firewall rules for each interface in a separate sheet, with aliases annotated in the source, destination, and port fields as comments.

### Example

The script will take an OPNsense XML configuration file, extract rules for each network interface, and write them to an Excel file with the following format:

| Aktiviert? | Action | Protocol | Source         | Port | Destination    | Port | Gateway | Schedule | Number of Rule | Description | Rules |
|------------|--------|----------|----------------|------|----------------|------|---------|----------|----------------|-------------|-------|
| Yes        | pass   | TCP      | 192.168.1.0/24 | 80   | 10.0.0.0/24    | 443  | *       | *        | 12345          | Allow Web   | *     |

- If `Source`, `Destination`, or `Port` fields contain an alias, you can hover over the cell to see the details of the alias.

### How It Works

- **Parse the XML**: The script uses Python’s `xml.etree.ElementTree` to parse the OPNsense firewall configuration file.
- **Extract Aliases**: The script identifies and extracts all aliases from the configuration and stores them for later use.
- **Add Comments**: When an alias is used in a source, destination, or port field, a comment with the alias details (such as IP ranges or ports) is added to the corresponding cell in the Excel sheet.

### Future Improvements

- Add support for custom Excel formatting or styles.
- Enhance error handling and support for more firewall rule fields.
- Support for exporting additional types of firewall configurations.

### Contributing

Contributions are welcome! If you find any issues or have ideas for improvements, feel free to open an issue or submit a pull request.
