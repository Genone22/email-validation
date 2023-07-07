import pandas as pd
import smtplib
import dns.resolver
import concurrent.futures
from tqdm import tqdm, auto

def validate_email(email):
    if '@' not in email:
        return False

    domain = email.split('@')[1]

    try:
        records = dns.resolver.resolve(domain, 'MX')
        mx_record = str(records[0].exchange)

        server = smtplib.SMTP()
        server.set_debuglevel(0)
        server.connect(mx_record)
        server.helo(server.local_hostname)
        server.mail('me@domain.com')

        code, message = server.rcpt(str(email))
        server.quit()

        if code == 250:
            return True
        else:
            return False
    except dns.resolver.NXDOMAIN:
        return False
    except smtplib.SMTPConnectError:
        return False
    except smtplib.SMTPServerDisconnected:
        return False
    except smtplib.SMTPResponseException:
        return False
    except:
        return False

def validate_emails():
    # Specify the input file path
    input_filepath = r"C:\Users\admin\Downloads\email'ы для стоматологий.xlsx"

    # Specify the output file path
    output_filepath = r"C:\Users\admin\Downloads\email'ы для стоматологий(валидные).xlsx"

    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_filepath)

    total_count = len(df)
    valid_count = 0
    invalid_count = 0

    # Create a progress bar
    progress_bar = tqdm(auto.tqdm(total=total_count, desc="Progress", unit="email"))

    def validate_email_helper(row):
        email = row[0]
        if email and validate_email(email):
            return True
        else:
            return False

    # Process email validation in parallel using multiple threads
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(tqdm(executor.map(validate_email_helper, df.itertuples(index=False)), total=total_count))

    # Update the DataFrame with the validation results
    df['Valid'] = results

    # Write the validated email addresses to the output Excel file
    df[df['Valid']].to_excel(output_filepath, index=False)

    # Count valid and invalid email addresses
    valid_count = df['Valid'].sum()
    invalid_count = total_count - valid_count

    progress_bar.close()

    print("Total: {}".format(total_count))
    print("Valid: {}".format(valid_count))
    print("Invalid: {}".format(invalid_count))

if __name__ == '__main__':
    validate_emails()

