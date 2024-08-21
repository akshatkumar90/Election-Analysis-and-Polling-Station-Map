import os               # interact with local system
import requests         # for requests 
import time                # for sleep time 

# Suppress insecure request warnings (not recommended for production)
requests.packages.urllib3.disable_warnings()

def download_pdf(url, directory):
    filename = os.path.join(directory, url.split("/")[-1])
    with open(filename, 'wb') as f:
        response = requests.get(url, verify=False)  # Disable SSL certificate verification
        f.write(response.content)

def download_pdfs(base_url, year, directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

    for i in range(1, 289):
        url = f"{base_url}{year}/AC_{str(i).zfill(3)}.pdf"
        download_pdf(url, directory)
        print(f"Downloaded {url}")
        time.sleep(5)

def main():
    base_url = "https://ceoelection.maharashtra.gov.in/form20/"  # url
    years = ["AC2009","AC2014","AC2019"]                            # year list
    directory = r"C:\Users\Akshat\Desktop\Project 6S\Election2.0\PollingStationResult" # local machine file path 

    for year in years:
        download_pdfs(base_url, year, os.path.join(directory, year))

if __name__ == "__main__":
    main()
