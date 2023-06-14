import re
import yaml
from pathlib import Path
import argparse
import win32com.client
from datetime import datetime
from collections import OrderedDict

def load_config(file_path):
    """
    Load the configuration from a YAML file.

    Args:
        file_path (Path): Path to the YAML configuration file.

    Returns:
        dict: Dictionary containing the configuration options.
    """
    with open(file_path, 'r') as file:
        config = yaml.safe_load(file)
    return config

def load_regex_patterns(file_path):
    """
    Load patterns of regular expression from a YAML file.

    Args:
        file_path (Path): Path to the YAML file containing the regex patterns.

    Returns:
        dict: Dictionary containing the loaded regex patterns.
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        regex_patterns = yaml.safe_load(file)
    return regex_patterns

def extract_items(text, regex_patterns):
    """
    Extract items from the text based on the provided regex patterns.

    Args:
        text (str): Text to be processed.
        regex_patterns (dict): Dictionary containing the regex patterns for each item.

    Returns:
        dict: Dictionary containing the extracted items.
    """
    extracted_items = {}
    for item, pattern in regex_patterns.items():
        regex = re.compile(pattern)
        match = regex.search(text)
        if match:
            extracted_items[item] = match.group(1)
    return extracted_items

def read_outlook_emails(start_date, end_date):
    """
    Read Outlook emails within the specified date range.

    Args:
        start_date (datetime): Start date and time of the period to search for emails.
        end_date (datetime): End date and time of the period to search for emails.

    Returns:
        list: List of OrderedDict objects, each representing a single email with its properties.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.GetDefaultFolder(6)  # 6 represents the Inbox folder

    # Construct the filter based on start and end dates
    filter_str = "[ReceivedTime] >= '" + start_date.strftime('%m/%d/%Y %H:%M %p') + "'"
    filter_str += " AND [ReceivedTime] <= '" + end_date.strftime('%m/%d/%Y %H:%M %p') + "'"

    # Search for emails within the specified date range
    items = folder.Items.Restrict(filter_str)

    emails = []
    for item in items:
        email = OrderedDict()
        email['Subject'] = item.Subject
        email['Sender'] = item.SenderName
        email['ReceivedTime'] = item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
        email['Body'] = item.Body
        emails.append(email)

    return emails

def parse_arguments():
    """
    Parse command line arguments.

    Returns:
        argparse.Namespace: Parsed command line arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('--config-file', type=Path, help='Path to the YAML configuration file.')
    parser.add_argument('--regex-file', type=Path, help='Path to the YAML file containing the regex patterns.')
    return parser.parse_args()

def main():
    # Parse command line arguments
    args = parse_arguments()

    # Load regex patterns from the YAML file
    regex_patterns = load_regex_patterns(args.regex_file)

    # Load the configuration from the YAML file
    config = load_config(args.config_file)

    # Extract start and end dates from the configuration
    start_date = datetime.strptime(config['start_date'], '%Y-%m-%d %H:%M:%S')
    end_date = datetime.strptime(config['end_date'], '%Y-%m-%d %H:%M:%S')

    # Read Outlook emails within the specified date range
    emails = read_outlook_emails(start_date, end_date)

    # Print the email properties
    for email in emails:
        for key, value in email.items():
            print(f'{key}: {value}')
            if key == 'Body':
                extracted_items = extract_items(value, regex_patterns)
                for key2, value2 in extracted_items.items():
                    print(f'{key2}: {value2}')
                    
        print('--------------------------------')

if __name__ == '__main__':
    main()