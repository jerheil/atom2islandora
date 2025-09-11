import csv
import re
import zipfile
import os

def extract_and_rename_zip(source_dir):
    for fname in os.listdir(source_dir):
        if fname.lower().endswith('.zip'):
            zip_path = os.path.join(source_dir, fname)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(source_dir)
                for extracted_name in zip_ref.namelist():
                    if extracted_name.lower().endswith('.csv'):
                        src = os.path.join(source_dir, extracted_name)
                        dst = os.path.join(source_dir, 'source1.csv')
                        if os.path.basename(src) != 'source1.csv':
                            os.rename(src, dst)
                        else:
                            dst = src
                        print(f"Extracted and renamed {extracted_name} to source1.csv")
                        return dst
            break
    else:
        raise FileNotFoundError("No zip file found in the source directory.")
    raise FileNotFoundError("No CSV file found in the zip archive.")

def load_source2(source2_path):
    mapping = {}
    with open(source2_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            mapping[row['SourceFile']] = row
            mapping[row['FileName']] = row
    return mapping

def normalize_location(loc):
    return re.sub(r"[ .-]", "", loc).lower()

def normalize_sourcefile(sf):
    return re.sub(r"[.-]", "", sf).lower()

def sr_mi_dot_match(phys_obj_loc, source2_mapping):
    match = re.match(r"^(SR|MI)\s*(\d+)\.(\d+)", phys_obj_loc)
    if match:
        prefix = match.group(1)
        main_num = match.group(2)
        sub_num = match.group(3)
        dotted_loc = f"{prefix}{main_num}.{sub_num}".lower()
        for key in source2_mapping:
            mod_key = key
            if "_" in key:
                mod_key = key.replace("_", ".", 1)
            mod_key_base = mod_key.rsplit('.', 1)[0].lower()
            if mod_key_base.startswith(dotted_loc):
                return source2_mapping[key]
    return None

def find_best_source2_match(phys_obj_loc, source2_mapping):
    if (phys_obj_loc.startswith("SR ") or phys_obj_loc.startswith("MI ")) and "." in phys_obj_loc:
        srmi_match = sr_mi_dot_match(phys_obj_loc, source2_mapping)
        if srmi_match:
            return srmi_match
    norm_loc = normalize_location(phys_obj_loc)
    for key in source2_mapping:
        norm_key = normalize_sourcefile(key)
        if norm_loc == norm_key:
            return source2_mapping[key]
        if norm_loc.replace(" ", "") == norm_key:
            return source2_mapping[key]
        if norm_loc in norm_key or norm_key in norm_loc:
            return source2_mapping[key]
    return None

def reformat_extent_and_medium(extent_and_medium):
    if not extent_and_medium:
        return ""
    parts = [
        re.sub(r"^\*\s*", "", line).strip()
        for line in extent_and_medium.split("\n")
        if line.strip()
    ]
    return ', '.join(parts)

def get_model_and_resource_type(mime):
    if mime.startswith("audio/"):
        return ("Audio", "Sound")
    elif mime.startswith("video/"):
        return ("Video", "Moving Image")
    elif mime.startswith("image/"):
        return ("Image", "Image")
    elif mime == "application/pdf":
        return ("Digital Document", "Text")
    else:
        return ("Other", "Other")

def write_error_report(error_rows, blank_rows, error_path, product_header):
    with open(error_path, 'w', encoding='utf-8') as ef:
        ef.write("Rows from source1.csv that could not be matched to source2.csv:\n")
        for err in error_rows:
            ef.write(f"Row {err['rownum']}: referenceCode={err.get('referenceCode', '')}, physicalObjectLocation={err.get('physicalObjectLocation', '')}, title={err.get('title','')}\n")
        ef.write("\nRows in product.csv with blank fields:\n")
        for blank in blank_rows:
            ef.write(f"Row {blank['rownum']} is missing fields: {', '.join(blank['fields'])}\n")
            ef.write(f"Values: {dict(zip(product_header, blank['values']))}\n")

def source1_to_product(source1_path, source2_mapping, output_path, member_of_existing_entity_id, error_path):
    header = [
        'ID', 'member_of_existing_entity_id', 'model', 'digital_file', 'mime', 'title', 'resource_type',
        'language', 'local_identifier', 'persons', 'description', 'origin_information', 'extent',
        'physical_location', 'shelf_locator'
    ]
    error_rows = []
    blank_rows = []
    with open(source1_path, newline='', encoding='utf-8') as f1, \
         open(output_path, 'w', newline='', encoding='utf-8') as outf:

        reader = csv.DictReader(f1)
        writer = csv.writer(outf)
        writer.writerow(header)
        id_counter = 1
        for rownum, row in enumerate(reader, 2):  # 2: header is line 1
            phys_obj_loc = row.get('physicalObjectLocation', '')
            source2_row = find_best_source2_match(phys_obj_loc, source2_mapping)
            if not source2_row:
                err = {
                    'rownum': rownum,
                    'referenceCode': row.get('referenceCode', ''),
                    'physicalObjectLocation': phys_obj_loc,
                    'title': row.get('title', '')
                }
                error_rows.append(err)
                continue

            mime = source2_row['MIMEType']
            model, resource_type = get_model_and_resource_type(mime)
            digital_file = f"repo-ingest://archives/{source2_row['FileName']}"
            title = row.get('title', '')
            language = 'English'
            local_identifier = row.get('referenceCode', '')
            event_actors = row.get('eventActors', '').strip()
            other_persons = (
                row.get('radTitleStatementOfResponsibility', '') or
                row.get('radTitleStatementOfResponsibilityNote', '') or
                row.get('radTitleAttributionsAndConjectures', '') or
                row.get('radNoteAccompanyingMaterial', '') or
                ''
            ).strip()
            if event_actors and other_persons:
                persons = f"{event_actors}; {other_persons}"
            else:
                persons = event_actors or other_persons

            description = row.get('scopeAndContent', '')
            origin_information = f"{row.get('eventTypes', '')}||{row.get('eventDates', '')}"
            extent = reformat_extent_and_medium(row.get('extentAndMedium', ''))
            physical_location = row.get('repository', '')
            shelf_locator = row.get('physicalObjectName', '')

            out_row = [
                str(id_counter),
                member_of_existing_entity_id,
                model,
                digital_file,
                mime,
                title,
                resource_type,
                language,
                local_identifier,
                persons,
                description,
                origin_information,
                extent,
                physical_location,
                shelf_locator
            ]
            # Check for blanks
            blank_fields = [header[i] for i, val in enumerate(out_row) if val == ""]
            if blank_fields:
                blank_rows.append({'rownum': rownum, 'fields': blank_fields, 'values': out_row})

            writer.writerow(out_row)
            id_counter += 1

    write_error_report(error_rows, blank_rows, error_path, header)

if __name__ == "__main__":
    source_dir = "."  # Change this if needed
    source1_path = extract_and_rename_zip(source_dir)
    source2_path = os.path.join(source_dir, "source2.csv")
    output_path = os.path.join(source_dir, "product.csv")
    error_path = os.path.join(source_dir, "error.txt")
    member_of_existing_entity_id = input("Enter value for member_of_existing_entity_id: ").strip()
    source2_mapping = load_source2(source2_path)
    source1_to_product(source1_path, source2_mapping, output_path, member_of_existing_entity_id, error_path)
    print("product.csv generated successfully.")
    print("error.txt written for unmatched rows and blank fields.")