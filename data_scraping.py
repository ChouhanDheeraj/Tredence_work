import requests
from bs4 import BeautifulSoup
import pandas as pd
from faker import Faker
import random
import great_expectations as gx
from great_expectations.core import ExpectationSuite, ExpectationConfiguration
from great_expectations.core.batch import RuntimeBatchRequest
from great_expectations.exceptions import DataContextError # Already there, ensure it is

# This is my test branch for scraping FHIR resources
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
# patient_url = "https://www.hl7.org/fhir/patient.html"
# df_patient = scrape_fhir_resource(patient_url, "Patient")

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

# Define file paths
SCHEMA_FILE = "FHIR_All_Resources_Schema_1.xlsx"
SYNTHETIC_DATA_FILE = "Synthetic_FHIR_Data.xlsx"

def initialize_gx_context():
    """
    Initializes or loads a Great Expectations Data Context.
    """
    import os
    try:
        # gx.get_context() will create the directory and config if it doesn't exist,
        # or load it if it does.
        context = gx.get_context(project_root_dir='.')
        # Verify that the directory and yml file were indeed created/exist.
        # GX 1.x creates a `gx` directory by default.
        if os.path.exists(os.path.join('.', 'gx', 'great_expectations.yml')):
            print("✅ Great Expectations context loaded/initialized successfully at './gx'.")
        else:
            # This is a fallback message if the directory/file isn't there,
            # which would be unexpected if get_context works as documented.
            print("⚠️ Great Expectations context obtained, but config file at './gx/great_expectations.yml' not found. Manual check might be needed.")
    except Exception as e:
        # Catch any other unexpected errors during context initialization.
        print(f"❌ An error occurred during Great Expectations context initialization: {e}")
        print("Please ensure Great Expectations is installed correctly and the project structure is as expected.")
        # Return None or raise if context cannot be established
        return None
    return context

def ensure_schema_file_exists():
    """
    Ensures the FHIR schema Excel file exists, creating it if necessary.
    """
    try:
        # Try to read the file to see if it exists and is valid
        pd.read_excel(SCHEMA_FILE, sheet_name=None) # Reading all sheets to check validity
        print(f"✅ Schema file '{SCHEMA_FILE}' already exists.")
    except FileNotFoundError:
        print(f"📋 Schema file '{SCHEMA_FILE}' not found. Scraping schemas...")
        # Save each resource DataFrame to a separate sheet
        with pd.ExcelWriter(SCHEMA_FILE, engine="openpyxl") as writer:
            for name, url in resource_urls.items():
                print(f"🔍 Scraping {name}...")
                df = scrape_fhir_resource(url, name)
                if not df.empty:
                    df.to_excel(writer, sheet_name=name[:31], index=False)
                else:
                    print(f"⚠️ Skipped {name} due to missing data.")
        print(f"✅ Schema file '{SCHEMA_FILE}' created successfully.")


def generate_synthetic_data(schema_df, num_records):
    """
    Generates synthetic data based on a schema DataFrame.
    For now, this is a placeholder and will be implemented in detail later.
    """
    """
    Generates synthetic data based on the provided schema.

    Args:
        schema_df (pd.DataFrame): DataFrame containing the schema for a resource.
        num_records (int): The number of records to generate.

    Returns:
        pd.DataFrame: DataFrame containing the generated synthetic data.
    """
    '''
    Generates synthetic data based on a schema DataFrame.
    '''
    fake = Faker()
    all_records = []
    for _ in range(num_records):
        record = {}
        for _, row in schema_df.iterrows():
            element_name = row['columns']
            # Ensure 'description' (datatype) and 'datatype' (cardinality) columns exist
            if 'description' not in row or 'datatype' not in row:
                # Handle missing schema columns if necessary, e.g., skip or log
                record[element_name] = "Error: Missing schema info"
                continue

            fhir_datatype = row['description']
            fhir_cardinality = row['datatype']

            faker_func = get_faker_provider(fhir_datatype, fhir_cardinality, element_name, fake)

            generated_value = None
            if fhir_cardinality.endswith("..*") or fhir_cardinality.endswith("..* "): # Handles "0..*" and "1..*"
                min_items = 0
                if fhir_cardinality.startswith("1"):
                    min_items = 1

                num_items = random.randint(min_items, min_items + 2) # Generate 0-2 for 0..* and 1-3 for 1..*

                if num_items == 0 and fhir_cardinality.startswith("0"): # For 0..* specifically
                    generated_value = [] # or None, depending on how you want to represent empty list for optional repeated
                else:
                    temp_list = []
                    for _ in range(num_items):
                        item_value = faker_func()
                        # If faker_func is wrapped by optional_field, item_value can be None
                        # Do not add None to list if it's for a repeated element, unless the element itself is a list of optional items (complex case)
                        # For now, filter out None values from lists, unless list itself can be None for 0..*
                        if item_value is not None:
                            temp_list.append(item_value)

                    if not temp_list and fhir_cardinality.startswith("0"): # if list is empty and it's 0..*
                        generated_value = [] # or None
                    elif not temp_list and fhir_cardinality.startswith("1"): # if list is empty but it's 1..* (should not happen if min_items=1 and faker_func produces non-None)
                        # This case indicates an issue or that faker_func returned None for a mandatory item.
                        # Forcing a single item generation if list is empty for 1..*
                        single_item = faker_func()
                        while single_item is None: # Ensure mandatory item is not None
                            single_item = get_faker_provider(fhir_datatype, "1..1", element_name, fake)() # Force non-optional
                        generated_value = [single_item]
                    else:
                        generated_value = temp_list

            elif fhir_cardinality == "0..1": # Optional single item
                generated_value = faker_func() # faker_func is already wrapped with optional_field logic
            elif fhir_cardinality == "1..1": # Mandatory single item
                # Ensure the value is not None for 1..1 fields
                generated_value = faker_func()
                while generated_value is None: # Should not be None if optional_field is correctly bypassed
                    # Re-fetch provider without optional wrapper if it somehow got one.
                    # This implies get_faker_provider needs to be careful for "1..1"
                    non_optional_provider = get_faker_provider(fhir_datatype, "1..1", element_name, fake)
                    generated_value = non_optional_provider()
            else: # Default or unknown cardinality
                generated_value = f"TODO_cardinality: {fhir_cardinality}"

            record[element_name] = generated_value
        all_records.append(record)

    return pd.DataFrame(all_records)

def get_faker_provider(fhir_datatype_str, fhir_cardinality_str, element_name_str, fake: Faker):
    '''
    Chooses the correct Faker method based on FHIR data type, cardinality, and element name.
    '''
    def optional_field(provider_func, null_probability=0.1):
        if random.random() < null_probability:
            return None
        return provider_func()

    is_optional = fhir_cardinality_str.startswith("0")
    provider = None

    # Enhanced type mappings
    if fhir_datatype_str == "string":
        if "description" in element_name_str.lower() or "text" in element_name_str.lower() or "note" in element_name_str.lower():
            provider = fake.text
        elif "name" in element_name_str.lower() and "given" not in element_name_str.lower() and "family" not in element_name_str.lower(): # Avoid conflict with HumanName parts
            provider = fake.name # For general names, not person names
        else:
            provider = fake.sentence # Default for other strings
    elif fhir_datatype_str == "uri" or fhir_datatype_str == "url":
        provider = fake.uri
    elif fhir_datatype_str == "boolean":
        provider = fake.boolean
    elif fhir_datatype_str == "date":
        if "birthDate" in element_name_str:
            provider = fake.date_of_birth
        else:
            provider = fake.date_this_decade # More relevant dates
    elif fhir_datatype_str == "dateTime":
        provider = fake.date_time_this_decade # More relevant datetimes
    elif fhir_datatype_str == "instant":
        provider = fake.iso8601 # Represents a point in time
    elif fhir_datatype_str == "integer" or fhir_datatype_str == "positiveInt" or fhir_datatype_str == "unsignedInt":
        min_val = 0 if fhir_datatype_str != "integer" else -100 # Allow negative for general integer
        max_val = 1000 if fhir_datatype_str != "integer" else 100
        provider = lambda: fake.random_int(min=min_val, max=max_val)
    elif fhir_datatype_str == "decimal":
        provider = lambda: fake.pydecimal(left_digits=random.randint(1,5), right_digits=random.randint(0,4), positive=True)
    elif fhir_datatype_str == "HumanName":
        provider = fake.name # Faker's name can be used directly
    elif fhir_datatype_str == "Address":
        provider = fake.address
    elif fhir_datatype_str == "ContactPoint":
        # This is a complex type, often needs specific handling for system (phone, email, etc.)
        # For now, a generic approach or more specific based on element_name
        if "phone" in element_name_str.lower():
            provider = fake.phone_number
        elif "email" in element_name_str.lower():
            provider = fake.email
        else: # Fallback for other contact points like 'fax', 'pager', 'url'
            provider = fake.word
    elif fhir_datatype_str == "Identifier":
        # Identifiers can be diverse. Using UUID as a common placeholder.
        # Real identifiers often have specific formats (e.g., MRN, SSN).
        provider = fake.uuid4
    elif fhir_datatype_str == "code": # simple code, often bound to a value set
        provider = lambda: fake.lexify(text="????", letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") # Generic code
    elif fhir_datatype_str == "id": # Resource id
        provider = lambda: fake.uuid4().split('-')[0] # Short id
    elif fhir_datatype_str == "CodeableConcept":
        # Complex type. For now, generate a plausible text description.
        # Real implementation might involve picking from a set of codes/texts.
        provider = fake.bs
    elif fhir_datatype_str == "Coding":
        # Complex type, part of CodeableConcept. Generate plausible code and display.
        provider = lambda: {"system": fake.uri(), "code": fake.lexify("????"), "display": fake.sentence()}
    elif fhir_datatype_str == "Period":
        provider = lambda: {"start": fake.date_time_this_decade().isoformat(), "end": fake.date_time_this_decade().isoformat()}
    elif fhir_datatype_str == "Quantity":
        provider = lambda: {"value": fake.pydecimal(left_digits=3, right_digits=2, positive=True), "unit": fake.word()}
    elif fhir_datatype_str == "Range":
        provider = lambda: {"low": {"value": fake.random_int(min=0, max=50)}, "high": {"value": fake.random_int(min=51, max=100)}}
    elif fhir_datatype_str == "Reference": # Represents a reference to another resource
        # For now, a simple string placeholder. Real references need careful construction.
        provider = lambda: f"{fake.word().capitalize()}/{fake.uuid4()}"
    elif fhir_datatype_str == "Narrative":
        provider = lambda: {"status": "generated", "div": f"<div xmlns='http://www.w3.org/1999/xhtml'>{fake.paragraph()}</div>"}
    elif fhir_datatype_str == "Extension":
        provider = lambda: {"url": fake.uri(), "valueString": fake.sentence()} # Simplified Extension
    else:
        # Fallback for unmapped types
        provider = lambda: f"TODO_unmapped_type: {fhir_datatype_str}"

    if is_optional:
        if not fhir_cardinality_str.endswith("*"): # Handles 0..1
            return lambda: optional_field(provider, null_probability=0.5)
        else: # Handles items within a 0..* list, each item can be None via optional_field
            return lambda: optional_field(provider, null_probability=0.1)
    else: # Mandatory field
        return provider

def get_or_create_expectation_suite(context: gx.DataContext, suite_name: str, schema_df: pd.DataFrame):
    '''
    Gets an existing Expectation Suite or creates a new one with basic expectations
    derived from the schema_df.
    '''
    try:
        suite = context.get_expectation_suite(suite_name)
        print(f"Found existing suite '{suite_name}'.")
    except gx.exceptions.GreatExpectationsError: # More specific exception if possible, e.g. SuiteNotFoundError
        print(f"Suite '{suite_name}' not found. Creating a new one.")
        suite = context.add_expectation_suite(suite_name)

    # Basic expectations based on schema
    for _, row in schema_df.iterrows():
        col_name = row['columns']
        cardinality = str(row['datatype']) # Ensure it's a string
        fhir_type = str(row['description']) # Ensure it's a string

        # Expect column to exist
        suite.add_expectation(ExpectationConfiguration(
            expectation_type="expect_column_to_exist",
            kwargs={"column": col_name}
        ))

        # Expect column values to not be null for mandatory fields
        if cardinality.startswith("1"): # e.g., "1..1" or "1..*"
            suite.add_expectation(ExpectationConfiguration(
                expectation_type="expect_column_values_to_not_be_null",
                kwargs={"column": col_name}
            ))

        # Basic type checking (can be expanded)
        # This is a simplified mapping. Real FHIR types are more complex.
        python_type = None
        if fhir_type.lower() in ["boolean"]:
            python_type = "bool"
        elif fhir_type.lower() in ["integer", "positiveint", "unsignedint"]:
            python_type = "int" # or "int64" for pandas
        elif fhir_type.lower() in ["decimal"]:
            python_type = "float" # or "float64" for pandas, or "Decimal" if using decimal objects
        elif fhir_type.lower() in ["string", "uri", "url", "code", "id", "markdown", "base64binary", "oid", "uuid", "xhtml"]:
            python_type = "str"
        elif fhir_type.lower() in ["date", "datetime", "instant", "time"]:
            # FordateTime and date, could also check format if values are strings
            python_type = "str" # Assuming they are stored as strings first
        # Add more type mappings as needed for types like HumanName, Address, etc.
        # These often are dicts or lists of dicts if complex.
        # For now, we'll only add type expectations for simple primitive-like types.

        if python_type:
            suite.add_expectation(ExpectationConfiguration(
                expectation_type="expect_column_values_to_be_of_type",
                kwargs={"column": col_name, "type_": python_type}
            ))
            # For strings, you might also check for empty strings if not allowed by cardinality
            if python_type == "str" and cardinality.startswith("1"):
                suite.add_expectation(ExpectationConfiguration(
                    expectation_type="expect_column_values_to_not_be_null", # Already added, but also means not empty string for some definitions
                    kwargs={"column": col_name, "mostly": 0.95} # Allow some flexibility for complex fields just stored as string
                ))

    context.save_expectation_suite(suite)
    print(f"Expectation suite '{suite_name}' saved with basic expectations.")
    return suite

def validate_dataframe(context: gx.DataContext, df_to_validate: pd.DataFrame, suite_name: str, resource_name: str):
    '''
    Validates a DataFrame against a given Expectation Suite using a RuntimeBatchRequest.
    '''
    datasource_name = "runtime_pandas_datasource" # Choose a name
    try:
        # Check if datasource exists, if not add it.
        # This is a simple way to handle runtime data.
        # For more permanent setup, configure in yml.
        context.get_datasource(datasource_name)
        print(f"Datasource '{datasource_name}' already exists.")
    except gx.exceptions.DatasourceNotFoundError:
        print(f"Datasource '{datasource_name}' not found. Adding a new one.")
        # Add a runtime datasource (Pandas)
        datasource_config = {
            "name": datasource_name,
            "class_name": "Datasource",
            "module_name": "great_expectations.datasource",
            "execution_engine": {
                "module_name": "great_expectations.execution_engine",
                "class_name": "PandasExecutionEngine",
            },
            "data_connectors": {
                f"runtime_data_connector_{resource_name}": { # Unique connector name per resource
                    "class_name": "RuntimeDataConnector",
                    "module_name": "great_expectations.datasource.data_connector",
                    "batch_identifiers": ["batch_id"], # Or any other identifiers
                }
            },
        }
        context.add_datasource(**datasource_config)
        print(f"Datasource '{datasource_name}' added.")

    batch_request = RuntimeBatchRequest(
        datasource_name=datasource_name,
        data_connector_name=f"runtime_data_connector_{resource_name}",
        data_asset_name=f"{resource_name}_asset", # Name for this data asset
        runtime_parameters={"batch_data": df_to_validate},
        batch_identifiers={"batch_id": f"batch_{resource_name}_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}"}
    )

    checkpoint_name = f"{resource_name}_checkpoint"
    try:
        # Try to get an existing checkpoint
        checkpoint = context.get_checkpoint(checkpoint_name)
        print(f"Found existing checkpoint '{checkpoint_name}'.")
    except gx.exceptions.CheckpointNotFoundError:
        print(f"Checkpoint '{checkpoint_name}' not found. Creating a new one.")
        checkpoint_config = {
            "name": checkpoint_name,
            "config_version": 1.0,
            "class_name": "SimpleCheckpoint",
            "run_name_template": f"%Y%m%d-%H%M%S-{resource_name}-validation",
            "validations": [
                {
                    "batch_request": batch_request, # This might need to be defined more generically if checkpoint is saved
                    "expectation_suite_name": suite_name,
                }
            ],
        }
        # A workaround for using RuntimeBatchRequest in a saved checkpoint is to not save it,
        # or use a Checkpoint that is configured to accept a batch_request at runtime.
        # For simplicity here, we'll run it directly without saving if it involves runtime batch request this way.
        # Or, we create a checkpoint that expects parameters.
        # Let's try to run an ad-hoc checkpoint (SimpleCheckpoint)
        print(f"Creating and running ad-hoc checkpoint for '{suite_name}'.")

    # Run the checkpoint
    # For SimpleCheckpoint, it's better to run it directly with parameters
    # If checkpoint was loaded and is not SimpleCheckpoint, its run method might differ
    results = context.run_checkpoint(
        checkpoint_name=None, # Run ad-hoc
        batch_request=batch_request,
        expectation_suite_name=suite_name,
        # action_list can be configured to store results, update data docs, etc.
        # Default actions are usually sufficient for validation results.
    )

    # Optional: Save validation results (if store is configured)
    # context.save_validation_results(validation_result_suite=results.list_validation_results()[0])

    print(f"Validation results for {resource_name}: Success = {results['success']}")
    if not results["success"]:
        print("Failed expectations:")
        for run_result in results["run_results"].values(): # Iterate through BatchResult objects
            for validation_result in run_result["validation_result"]["results"]: # list of ExpectationSuiteValidationResult
                if not validation_result["success"]:
                    print(f"  - Expectation: {validation_result['expectation_config']['expectation_type']}")
                    print(f"    Column: {validation_result['expectation_config']['kwargs'].get('column')}")
                    print(f"    Details: {validation_result.get('result', {})}")
    return results["success"]


if __name__ == "__main__":
    gx_context = initialize_gx_context() # Keep this
    ensure_schema_file_exists()

    try:
        excel_file = pd.ExcelFile(SCHEMA_FILE)
        resource_names = excel_file.sheet_names
        print(f"Found schemas for: {resource_names}")

        with pd.ExcelWriter(SYNTHETIC_DATA_FILE, engine="openpyxl") as writer:
            for resource_name in resource_names:
                print(f"\n🔄 Processing schema for {resource_name}...")
                current_schema_df = pd.read_excel(excel_file, sheet_name=resource_name)

                if not all(col in current_schema_df.columns for col in ['columns', 'description', 'datatype']):
                    print(f"⚠️ Schema for {resource_name} is missing required columns. Skipping.")
                    continue

                print(f"🧬 Generating 10 synthetic records for {resource_name}...")
                synthetic_df = generate_synthetic_data(current_schema_df, num_records=10)

                if not synthetic_df.empty:
                    print(f"✔️ Synthetic data generated for {resource_name}. Proceeding to validation.")

                    # --- BEGIN GX INTEGRATION ---
                    current_suite_name = f"{resource_name}_suite"
                    get_or_create_expectation_suite(gx_context, current_suite_name, current_schema_df)

                    validation_passed = validate_dataframe(gx_context, synthetic_df, current_suite_name, resource_name)

                    if validation_passed:
                        print(f"✅ GX validation passed for {resource_name}.")
                        synthetic_df.to_excel(writer, sheet_name=resource_name[:31], index=False)
                        print(f"💾 Synthetic data for {resource_name} saved to sheet.")
                    else:
                        print(f"❌ GX validation failed for {resource_name}. Data not saved. See logs for details.")
                        # Optionally, here you could trigger regeneration or other error handling
                        # For now, we just skip saving if validation fails.
                    # --- END GX INTEGRATION ---

                else:
                    print(f"⚠️ No synthetic data generated for {resource_name} (DataFrame was empty).")

        print(f"\n🎉 All synthetic data saved to '{SYNTHETIC_DATA_FILE}' (if validation passed).")

    except FileNotFoundError:
        print(f"❌ Error: Schema file '{SCHEMA_FILE}' not found. Cannot generate synthetic data.")
    except Exception as e:
        print(f"An error occurred during synthetic data generation or validation: {e}")