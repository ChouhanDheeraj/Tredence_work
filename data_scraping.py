import requests
from bs4 import BeautifulSoup
import pandas as pd

def scrape_fhir_resource(url, resource_name):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find the StructureDefinition table
    table = soup.find('table', {'class': 'list'})


    rows = table.find_all('tr')[1:]  # Skip header
    data = []

    for row in rows:
        cols = row.find_all('td')
        if len(cols) >= 4:
            element = cols[0].text.strip()
            cardinality = cols[1].text.strip()
            datatype = cols[2].text.strip()
            description = cols[3].text.strip()

            data.append({
                'Table_name': resource_name,
                'columns': element,
                'datatype': cardinality,
                'description': datatype,
                'full_column_with_schema': description
            })

    return pd.DataFrame(data)

# Example usage
patient_url = "https://www.hl7.org/fhir/patient.html"
df_patient = scrape_fhir_resource(patient_url, "Patient")

# Show the top 5 rows
#print(df_patient.head())

# Save to CSV/Excel
#df_patient.to_excel("FHIR_Patient_Schema.xlsx", index=False)


resource_urls = {
    "Patient": "https://www.hl7.org/fhir/patient.html",
    "Encounter": "https://www.hl7.org/fhir/encounter.html",
    "Practitioner": "https://www.hl7.org/fhir/practitioner.html",
    "Condition": "https://www.hl7.org/fhir/condition.html"
}


# Save each resource DataFrame to a separate sheet
with pd.ExcelWriter("FHIR_All_Resources_Schema_1.xlsx", engine="openpyxl") as writer:
    for name, url in resource_urls.items():
        print(f"🔍 Scraping {name}...")
        df = scrape_fhir_resource(url, name)
        if not df.empty:
            df.to_excel(writer, sheet_name=name[:31], index=False)  # Excel sheet name max length = 31
        else:
            print(f"⚠️ Skipped {name} due to missing data.")