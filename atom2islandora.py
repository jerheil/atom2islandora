import csv
import re
import os
import subprocess
import tkinter as tk
from tkinter import filedialog
import zipfile

def select_folder_dialog(title="Select folder"):
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_selected

def run_exiftool_and_create_source2_csv(dest_folder, source_folder):
    exiftool_path = r"C:\Windows\exiftool.exe"
    output_csv = os.path.join(dest_folder, "source2.csv")
    args = [
        exiftool_path,
        "-csv", "-r",
        "-SourceFile", "-Title", "-FileName", "-FileCreateDate", "-PageCount",
        "-FileTypeExtension", "-MIMEType", "-LayerCount",
        "*"
    ]
    with open(output_csv, "w", encoding="utf-8") as outfile:
        subprocess.run(args, cwd=source_folder, stdout=outfile, check=True)
    print(f"source2.csv generated at {output_csv}")

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
    # Expanded logic: any "_" between two digits is treated as "."
    mapping = {}
    with open(source2_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            for keyfield in ["SourceFile", "FileName"]:
                key = row.get(keyfield, "")
                if key:
                    mapping[key] = row
                    # Add dot-mapped version if applicable (e.g. MI12_3.tif --> MI12.3.tif)
                    dot_key = re.sub(r'(?<=\d)_(?=\d)', '.', key)
                    if dot_key != key:
                        mapping[dot_key] = row
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
            mod_key = re.sub(r'(?<=\d)_(?=\d)', '.', key)
            mod_key_base = mod_key.rsplit('.', 1)[0].lower()
            if mod_key_base.startswith(dotted_loc):
                return source2_mapping[key]
    return None

def normalize_audio_shelf(shelf_locator):
    # Normalize: remove spaces and dots, replace dot with underscore
    # E.g., "SR 1267.435" -> "SR1267_435"
    return shelf_locator.replace(" ", "").replace(".", "_")

def extract_side_label(sourcefile, normalized_audio_prefix):
    # Extract SideLabel, e.g. "SideA", "SideB", "Side1", "Side2", "Tape1" etc. from SourceFile after the shelf locator
    # Example: SR1267_435-PhilBrown-CFRCReminiscences1-SideA.mp3 -> "SideA"
    if not sourcefile:
        return ""
    base = os.path.splitext(os.path.basename(sourcefile))[0]
    # attempt pattern: prefix-...-SideA
    pattern = rf"{re.escape(normalized_audio_prefix)}[-_][^-_]*[-_](Side[A-Za-z0-9]+)$"
    match = re.search(pattern, base, re.IGNORECASE)
    if match:
        return match.group(1)
    # fallback: common endings
    for label in ["SideA", "SideB", "Side1", "Side2"] + [f"Tape{i}" for i in range(1, 11)]:
        if base.lower().endswith(label.lower()):
            return base[-len(label):]
    # final fallback: return substring after prefix if present
    pref_pattern = rf"^{re.escape(normalized_audio_prefix)}[-_]*"
    after_prefix = re.sub(pref_pattern, "", base, flags=re.IGNORECASE)
    if after_prefix:
        after_prefix = after_prefix.lstrip("-_")
        return after_prefix
    return ""

def is_compound(row, source2_mapping):
    """
    Detect image/audio children for a given source1 row by scanning source2 mapping.
    Returns (image_children_list, audio_children_list).
    """
    phys_obj_loc = (row.get('physicalObjectLocation') or '').strip()
    shelf_locator = (row.get('physicalObjectLocation') or row.get('shelf_locator') or row.get('shelfLocator') or '').strip()
    norm_phys_obj_loc = phys_obj_loc.replace(".", "_").replace(" ", "")
    norm_shelf_locator = normalize_audio_shelf(shelf_locator)

    image_children = []
    audio_children = []
    seen_rows = set()
    for s2key, s2row in source2_mapping.items():
        # s2key may be full path or filename; get basename without extension
        cand = os.path.splitext(os.path.basename(s2key))[0]
        if cand.lower().startswith((norm_phys_obj_loc + "-").lower()) or cand.lower().startswith((norm_phys_obj_loc + "_").lower()):
            if id(s2row) not in seen_rows:
                seen_rows.add(id(s2row))
                mt = s2row.get('MIMEType', '').lower()
                if mt.startswith('image/'):
                    image_children.append(s2row)
                elif mt.startswith('audio/'):
                    audio_children.append(s2row)
        elif cand.lower().startswith((norm_shelf_locator + "-").lower()) or cand.lower().startswith((norm_shelf_locator + "_").lower()):
            if id(s2row) not in seen_rows:
                seen_rows.add(id(s2row))
                mt = s2row.get('MIMEType', '').lower()
                if mt.startswith('image/'):
                    image_children.append(s2row)
                elif mt.startswith('audio/'):
                    audio_children.append(s2row)
    return (image_children, audio_children)

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
    if not mime:
        return ("Other", "Other")
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
            ef.write(f"Row {err.get('rownum','?')}: referenceCode={err.get('referenceCode', '')}, physicalObjectLocation={err.get('physicalObjectLocation', '')}, title={err.get('title','')}\n")
        ef.write("\nRows in product.csv with blank fields:\n")
        for blank in blank_rows:
            vals = dict(zip(product_header, blank['values']))
            member_of_existing_entity_id = vals.get('member_of_existing_entity_id', '')
            member_of = vals.get('member_of', '')
            # Only report if BOTH are blank, or BOTH are filled
            if (not member_of_existing_entity_id and not member_of) or (member_of_existing_entity_id and member_of):
                ef.write(f"Row {blank['rownum']} is missing fields: {', '.join(blank['fields'])}\n")
                ef.write(f"Values: {vals}\n")

def _normalize_for_matching(s):
    if not s:
        return []
    s = s.lower()
    variants = []
    no_space = re.sub(r"\s+", "", s)
    variants.append(no_space)
    variants.append(no_space.replace(".", "_"))
    variants.append(no_space.replace("_", "."))
    variants.append(re.sub(r"[._\-]", "", no_space))
    digits_dot = re.sub(r"(?<=\d)_(?=\d)", ".", no_space)
    variants.append(digits_dot)
    # unique preserve order
    uniq = []
    for v in variants:
        if v and v not in uniq:
            uniq.append(v)
    return uniq

def find_best_direct_matches(phys_obj_loc, shelf_locator, source2_mapping):
    """
    Broad matching logic:
    Returns (best_image_row, best_audio_row) or (None, None)
    """
    phys_variants = _normalize_for_matching(phys_obj_loc)
    shelf_variants = _normalize_for_matching(shelf_locator)

    # Unique rows to check
    rows_to_check = list({id(v): v for v in source2_mapping.values()}.values())

    best_image = None
    best_audio = None

    # helper to produce candidate basenames and compact form
    def cand_forms(row):
        candidates = []
        sf = row.get('SourceFile') or ""
        fn = row.get('FileName') or ""
        if sf:
            candidates.append(os.path.splitext(os.path.basename(sf))[0])
        if fn:
            candidates.append(os.path.splitext(os.path.basename(fn))[0])
        # dedupe
        seen = []
        for c in candidates:
            if c and c not in seen:
                seen.append(c)
        return seen

    # Pass 1: exact normalized equality
    for row in rows_to_check:
        for cand in cand_forms(row):
            cand_compact = re.sub(r"[._\-]", "", cand).lower()
            for var in phys_variants + shelf_variants:
                if re.sub(r"[._\-]", "", var).lower() == cand_compact:
                    mt = (row.get('MIMEType') or '').lower()
                    if mt.startswith('image/') and not best_image:
                        best_image = row
                    elif mt.startswith('audio/') and not best_audio:
                        best_audio = row
            if best_image and best_audio:
                return (best_image, best_audio)

    # Pass 2: startswith (permissive)
    for row in rows_to_check:
        for cand in cand_forms(row):
            cand_low = cand.lower()
            cand_compact = re.sub(r"[._\-]", "", cand_low)
            for var in phys_variants + shelf_variants:
                v = var.lower()
                v_compact = re.sub(r"[._\-]", "", v)
                if cand_low.startswith(v) or v.startswith(cand_low) or cand_compact.startswith(v_compact) or v_compact.startswith(cand_compact):
                    mt = (row.get('MIMEType') or '').lower()
                    if mt.startswith('image/') and not best_image:
                        best_image = row
                    elif mt.startswith('audio/') and not best_audio:
                        best_audio = row
            if best_image and best_audio:
                return (best_image, best_audio)

    # Pass 3: substring
    for row in rows_to_check:
        for cand in cand_forms(row):
            cand_low = cand.lower()
            cand_compact = re.sub(r"[._\-]", "", cand_low)
            for var in phys_variants + shelf_variants:
                v = var.lower()
                v_compact = re.sub(r"[._\-]", "", v)
                if v in cand_low or cand_low in v or v_compact in cand_compact or cand_compact in v_compact:
                    mt = (row.get('MIMEType') or '').lower()
                    if mt.startswith('image/') and not best_image:
                        best_image = row
                    elif mt.startswith('audio/') and not best_audio:
                        best_audio = row
        if best_image and best_audio:
            break

    return (best_image, best_audio)

def source1_to_product(source1_path, source2_mapping, output_path, member_of_existing_entity_id, error_path):
    import csv, re

    # Prompt for entity type
    is_corporate = False
    print("Are these records from a Corporation or Conceptual entity? (y/n)")
    answer = input().strip().lower()
    if answer == "y":
        is_corporate = True

    # Ask if Authorized form of name is the same for all records
    print(f"Is the Authorized form of name for {('organizations' if is_corporate else 'persons')} the same for all records? (y/n)")
    uniform_name_answer = input().strip().lower()
    uniform_authorized_name = ""
    if uniform_name_answer == "y":
        print(f"Please enter the Authorized form of name for all records ({'organizations' if is_corporate else 'persons'}):")
        uniform_authorized_name = input().strip()

    header = [
        'ID', 'member_of_existing_entity_id', 'member_of', 'model', 'digital_file', 'mime', 'title', 'resource_type',
        'language', 'local_identifier',
        'organizations' if is_corporate else 'persons',
        'description', 'origin_information', 'extent',
        'physical_location', 'shelf_locator', 'location_url'
    ]

    error_rows = []
    blank_rows = []
    collections_by_id = set()
    written_ids = set()
    next_compound_child_id = 1

    def get_next_child_id():
        nonlocal next_compound_child_id
        v = next_compound_child_id
        next_compound_child_id += 1
        return str(v)

    with open(source1_path, newline='', encoding='utf-8') as f1:
        reader = csv.DictReader(f1)
        source1_rows = list(reader)

    with open(output_path, 'w', newline='', encoding='utf-8') as outf:
        writer = csv.writer(outf)
        writer.writerow(header)

        # Pass 1: Write compound parents and their children
        for idx, row in enumerate(source1_rows):
            source1_id = row.get('legacyId', '').strip() or row.get('ID', '').strip() or str(idx + 1)
            phys_obj_loc = (row.get('physicalObjectLocation') or '').strip()
            shelf_locator = (row.get('physicalObjectLocation') or row.get('shelf_locator') or row.get('shelfLocator') or '').strip()
            title = (row.get('title') or '').strip()
            language = 'English'
            local_identifier = row.get('referenceCode', '')
            event_actors = (row.get('eventActors') or '').strip()
            other_persons = (
                row.get('radTitleStatementOfResponsibility', '') or
                row.get('radTitleStatementOfResponsibilityNote', '') or
                row.get('radTitleAttributionsAndConjectures', '') or
                row.get('radNoteAccompanyingMaterial', '') or
                ''
            ).strip()

            field_val = ""
            if event_actors and other_persons:
                field_val = f"{event_actors}; {other_persons}"
            else:
                field_val = event_actors or other_persons
            if not field_val or field_val.upper() == "NULL":
                field_val = uniform_authorized_name if uniform_authorized_name else field_val

            description = row.get('scopeAndContent', '')
            event_start = row.get('eventStartDates', '').strip()
            event_end = row.get('eventEndDates', '').strip()
            if event_start and event_end:
                if event_start == event_end:
                    origin_information = event_start
                else:
                    origin_information = f"{event_start}/{event_end}" if event_end else event_start
            elif event_start:
                origin_information = event_start
            elif event_end:
                origin_information = event_end
            else:
                origin_information = ""

            extent = reformat_extent_and_medium(row.get('extentAndMedium', ''))
            physical_location = row.get('repository', '')
            slug = row.get('slug', '').strip()
            location_url = f"https://db-archives.library.queensu.ca/{slug}" if slug else ""

            image_children, audio_children = is_compound(row, source2_mapping)

            if audio_children:
                written_ids.add(source1_id)
                compound_title = f"{title} - {shelf_locator}" if shelf_locator else title
                compound_row = [
                    source1_id, member_of_existing_entity_id, '', 'Compound', '', '', compound_title, 'Sound',
                    language, local_identifier, field_val, description, origin_information, extent, physical_location, shelf_locator, location_url
                ]
                writer.writerow(compound_row)
                for m in audio_children:
                    child_id = get_next_child_id()
                    norm_shelf_locator = normalize_audio_shelf(shelf_locator)
                    side_label = extract_side_label(m.get('SourceFile') or m.get('FileName',''), norm_shelf_locator)
                    child_title = f"{title} - {shelf_locator}-{side_label}" if side_label else f"{title} - {shelf_locator}"
                    child_row = [
                        child_id, '', source1_id, 'Audio',
                        f"repo-ingest://archives/{m.get('FileName', m.get('SourceFile',''))}", m.get('MIMEType', ''),
                        child_title, 'Sound', '', '', '', '', '', '', '', '', location_url
                    ]
                    writer.writerow(child_row)

            if image_children:
                written_ids.add(source1_id)
                compound_title = f"{title} - {phys_obj_loc}" if phys_obj_loc else title
                compound_row = [
                    source1_id, member_of_existing_entity_id, '', 'Compound', '', '', compound_title, 'Image',
                    language, local_identifier, field_val, description, origin_information, extent, physical_location, phys_obj_loc, location_url
                ]
                writer.writerow(compound_row)
                for m in image_children:
                    child_id = get_next_child_id()
                    child_title = f"{title} - {os.path.splitext(m.get('SourceFile',''))[0]}" if m.get('SourceFile') else f"{title}"
                    child_row = [
                        child_id, '', source1_id, 'Image', f"repo-ingest://archives/{m.get('FileName', m.get('SourceFile',''))}", m.get('MIMEType', ''),
                        child_title, 'Image', '', '', '', '', '', '', '', '', location_url
                    ]
                    writer.writerow(child_row)

        # Pass 2: write remaining individual items (non-compound)
        for idx, row in enumerate(source1_rows):
            source1_id = row.get('legacyId', '').strip() or row.get('ID', '').strip() or str(idx + 1)
            if source1_id in written_ids:
                continue
            phys_obj_loc = (row.get('physicalObjectLocation') or '').strip()
            shelf_locator = (row.get('physicalObjectLocation') or row.get('shelf_locator') or row.get('shelfLocator') or '').strip()
            title = (row.get('title') or '').strip()
            language = 'English'
            local_identifier = row.get('referenceCode', '')
            event_actors = (row.get('eventActors') or '').strip()
            other_persons = (
                row.get('radTitleStatementOfResponsibility', '') or
                row.get('radTitleStatementOfResponsibilityNote', '') or
                row.get('radTitleAttributionsAndConjectures', '') or
                row.get('radNoteAccompanyingMaterial', '') or
                ''
            ).strip()

            field_val = ""
            if event_actors and other_persons:
                field_val = f"{event_actors}; {other_persons}"
            else:
                field_val = event_actors or other_persons
            if not field_val or field_val.upper() == "NULL":
                field_val = uniform_authorized_name if uniform_authorized_name else field_val

            description = row.get('scopeAndContent', '')
            event_start = row.get('eventStartDates', '').strip()
            event_end = row.get('eventEndDates', '').strip()
            if event_start and event_end:
                if event_start == event_end:
                    origin_information = event_start
                else:
                    origin_information = f"{event_start}/{event_end}" if event_end else event_start
            elif event_start:
                origin_information = event_start
            elif event_end:
                origin_information = event_end
            else:
                origin_information = ""

            extent = reformat_extent_and_medium(row.get('extentAndMedium', ''))
            physical_location = row.get('repository', '')
            slug = row.get('slug', '').strip()
            location_url = f"https://db-archives.library.queensu.ca/{slug}" if slug else ""

            # Try to find direct image/audio matches using broadened matching
            direct_image, direct_audio = find_best_direct_matches(phys_obj_loc, shelf_locator, source2_mapping)

            if direct_audio:
                norm_shelf_locator = normalize_audio_shelf(shelf_locator)
                side_label = extract_side_label(direct_audio.get('SourceFile') or direct_audio.get('FileName',''), norm_shelf_locator)
                item_title = f"{title} - {shelf_locator}-{side_label}" if side_label else f"{title} - {shelf_locator}" if shelf_locator else title
                image_row = [
                    source1_id, member_of_existing_entity_id, '', 'Audio',
                    f"repo-ingest://archives/{direct_audio.get('FileName', direct_audio.get('SourceFile',''))}", direct_audio.get('MIMEType',''),
                    item_title, 'Sound', language, local_identifier, field_val, description, origin_information, extent,
                    physical_location, shelf_locator, location_url
                ]
                writer.writerow(image_row)
                continue

            if direct_image:
                sf_base = direct_image.get('SourceFile', '')
                sf_base_title = os.path.splitext(os.path.basename(sf_base))[0] if sf_base else ""
                item_title = f"{title} - {sf_base_title}" if sf_base_title else title
                image_row = [
                    source1_id, member_of_existing_entity_id, '', 'Image',
                    f"repo-ingest://archives/{direct_image.get('FileName', direct_image.get('SourceFile',''))}", direct_image.get('MIMEType',''),
                    item_title, 'Image', language, local_identifier, field_val, description, origin_information, extent,
                    physical_location, phys_obj_loc, location_url
                ]
                writer.writerow(image_row)
                continue

            # No direct match: write a normal row without digital_file
            image_row = [
                source1_id, member_of_existing_entity_id, '', 'Image',
                '', '', title, 'Image', language, local_identifier, field_val, description, origin_information, extent,
                physical_location, phys_obj_loc, location_url
            ]
            writer.writerow(image_row)

    # Write error/blank report if necessary (error_rows and blank_rows collected earlier if desired)
    write_error_report(error_rows, blank_rows, error_path, header)

def pad_photo_number(num, width=3):
    try:
        return str(int(num)).zfill(width)
    except Exception:
        return num

def parse_photo_numbers(photo_numbers):
    if not photo_numbers or not photo_numbers.strip():
        return []
    result = []
    parts = [x.strip() for x in photo_numbers.split(",")]
    for part in parts:
        if "-" in part:
            range_parts = [x.strip() for x in part.split("-")]
            if len(range_parts) == 2 and range_parts[0].isdigit() and range_parts[1].isdigit():
                start = int(range_parts[0])
                end = int(range_parts[1])
                if start <= end:
                    result.extend([str(n) for n in range(start, end + 1)])
                else:
                    result.extend([range_parts[0], range_parts[1]])
            else:
                result.append(part)
        else:
            if part:
                result.append(part)
    return [x for x in result if x]

def clean_fieldnames_and_rows(source1_path, cleaned_path):
    with open(source1_path, newline='', encoding='utf-8') as infile, \
         open(cleaned_path, 'w', newline='', encoding='utf-8') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)
        rows = list(reader)
        header = rows[0]

        # Deduplicate FLIGHT_LINE column names: keep first as 'FLIGHT_LINE', others renamed
        seen = {}
        clean_header = []
        for h in header:
            # Remove any extra characters before "Record_ID" (including spaces, punctuation, etc.)
            if "Record_ID" in h:
                clean_header.append("Record_ID")
                continue
            base = re.sub(r"\s.*", "", h.strip())
            if base not in seen:
                seen[base] = 1
                clean_header.append(base)
            else:
                # If duplicate, rename (e.g. FLIGHT_LINE_DUPLICATE)
                clean_header.append(f"{base}_DUPLICATE{seen[base]}")
                seen[base] += 1
        writer.writerow(clean_header)

        for row in rows[1:]:
            clean_row = []
            for v in row:
                v = v.lstrip()
                v = re.sub(r"^\s+$", " ", v)
                v = re.sub(r",\s+", ", ", v)
                clean_row.append(v)
            writer.writerow(clean_row)

def maps_mode_generate_product(source1_path, source2_path, output_path, mapping_report_path, member_of_existing_entity_id="10678"):
    cleaned_source1 = os.path.splitext(source1_path)[0] + "_cleaned.csv"
    clean_fieldnames_and_rows(source1_path, cleaned_source1)

    with open(cleaned_source1, newline='', encoding='utf-8') as f1:
        reader = csv.DictReader(f1)
        source1_rows = [dict(row) for row in reader]

    with open(source2_path, newline='', encoding='utf-8') as f2:
        reader = csv.DictReader(f2)
        # Map both SourceFile and FileName for robust lookup
        source2_rows = {}
        for row in reader:
            source2_rows[row.get('SourceFile', '')] = row
            if 'FileName' in row:
                source2_rows[row['FileName']] = row

    header = [
        "ID","local_item_identifier","local_identifier","title","physical_location",
        "hierarchical_geographic_subject","persons","origin_information","notes","description",
        "shelf_locator","resource_type","member_of_existing_entity_id","model","digital_file","mime"
    ]

    mapping_problems = []
    product_rows = []
    idx = 1
    for row in source1_rows:
        record_id = row.get("Record_ID", "")
        nts_map_no = row.get("NTS_MAP_NO", "")
        location = row.get("LOCATION", "")
        province = row.get("PROVINCE", "")
        persons = ""
        year = row.get("YEAR", "")
        scale = row.get("SCALE", "")
        notes = row.get("NOTES", "")
        shown = row.get("SHOWN", "")
        flight_line = row.get("FLIGHT_LINE", "")
        roll = row.get("ROLL", "")
        date = row.get("DATE", "")
        orientation = row.get("ORIENTATION", "")
        local_notes = row.get("LOCAL", "")
        photo_numbers_field = row.get("PHOTO_NUMBERS", "")
        image_link = row.get("IMAGE_LINK") or row.get("IMAGE") or ""

        photo_numbers = parse_photo_numbers(photo_numbers_field)
        if not photo_numbers:
            photo_numbers = [""]

        for photo_num in photo_numbers:
            # Title and shelf_locator mapping
            if location and flight_line and photo_num:
                if roll:
                    title = f'{location} (Flight Line {flight_line}, Roll [{roll}], Photo Number {photo_num})'
                    shelf_locator = f"Flight Line {flight_line}, Roll [{roll}], Photo Number {photo_num}"
                else:
                    title = f'{location} (Flight Line {flight_line}, Photo Number {photo_num})'
                    shelf_locator = f"Flight Line {flight_line}, Photo Number {photo_num}"
            elif location:
                title = location
                shelf_locator = flight_line or ""
            else:
                if flight_line and photo_num:
                    title = f"Flight Line {flight_line}, Photo Number {photo_num}"
                    shelf_locator = f"Flight Line {flight_line}, Photo Number {photo_num}"
                elif flight_line:
                    title = f"Flight Line {flight_line}"
                    shelf_locator = f"Flight Line {flight_line}"
                elif photo_num:
                    title = f"Photo Number {photo_num}"
                    shelf_locator = f"Photo Number {photo_num}"
                else:
                    title = ""
                    shelf_locator = ""

            physical_location = "Queen's University Maps and Air Photos Collection"
            hierarchical_geographic_subject = f"North America|Canada||{province}" if province else ""
            notes_field = f"scale|{scale}" if scale else ""
            description_pieces = []
            if date:
                description_pieces.append(date.strip())
            if orientation:
                description_pieces.append(orientation.strip())
            if notes:
                description_pieces.append(notes.strip())
            description = ". ".join(description_pieces)
            if description:
                description += "."
            elif shown:
                description = shown
            else:
                description = location

            resource_type = "Image"
            model = "Image"

            fl_for_file = flight_line.replace(" ", "")
            pn_padded = pad_photo_number(photo_num)
            filename_core = ""
            if fl_for_file and pn_padded:
                filename_core = f"{fl_for_file}_{pn_padded}.tif"
            elif image_link:
                filename_core = image_link
            else:
                filename_core = ""

            digital_filename = filename_core

            digital_file = f"repo-ingest://maps/{digital_filename}" if digital_filename else ""

            mime = ""
            if digital_filename and digital_filename in source2_rows and source2_rows[digital_filename].get("MIMEType"):
                mime = source2_rows[digital_filename]["MIMEType"]
            elif image_link and image_link in source2_rows and source2_rows[image_link].get("MIMEType"):
                mime = source2_rows[image_link]["MIMEType"]
            else:
                mapping_problems.append(
                    f"Could not find MIMEType for {digital_filename or image_link}"
                )

            member_id = member_of_existing_entity_id or ""

            product_row = [
                idx,                   # ID
                record_id,             # local_item_identifier
                nts_map_no,
                title,
                physical_location,
                hierarchical_geographic_subject,
                persons,
                year,
                notes_field,
                description,
                shelf_locator,
                resource_type,
                member_id,
                model,
                digital_file,
                mime
            ]
            product_rows.append(product_row)
            idx += 1

    with open(output_path, "w", newline='', encoding="utf-8") as fout:
        writer = csv.writer(fout)
        writer.writerow(header)
        for row in product_rows:
            writer.writerow(row)

    if mapping_problems:
        with open(mapping_report_path, "w", encoding="utf-8") as repf:
            repf.write("Mapping issues encountered during processing:\n")
            for mp in mapping_problems:
                repf.write(f"{mp}\n")
        print(f"Mapping issues written to {mapping_report_path}")
    else:
        print("No mapping issues encountered.")

    print(f"Map-mode product.csv generated at {output_path}")

def write_missing_metadata_report(product_path, source2_path, report_path):
    # Collect all digital filenames used in product.csv (these map to SourceFile in source2.csv)
    used_files = set()
    with open(product_path, newline='', encoding='utf-8') as pf:
        reader = csv.DictReader(pf)
        for row in reader:
            digital_file = row.get("digital_file", "") or row.get("digitalFile", "") or row.get("digital-file", "")
            # Extract filename (after last /)
            if digital_file:
                filename = digital_file.split("/")[-1]
                used_files.add(filename)

    # Collect all SourceFile from source2.csv
    missing_files = []
    with open(source2_path, newline='', encoding='utf-8') as s2f:
        reader = csv.DictReader(s2f)
        for row in reader:
            source_file = row.get("SourceFile", "")
            if source_file and source_file not in used_files:
                missing_files.append(source_file)

    # Write report
    with open(report_path, "w", encoding="utf-8") as rpt:
        if missing_files:
            rpt.write("The following SourceFile(s) in source2.csv did not match any digital_file in product.csv:\n")
            for fn in missing_files:
                rpt.write(f"{fn}\n")
        else:
            rpt.write("All SourceFile entries in source2.csv were matched in product.csv.\n")

if __name__ == "__main__":
    print("Are you creating this for Archives (a) or Map (m)?")
    mode = input("Type 'a' for Archives or 'm' for Map: ").strip().lower()
    if mode == 'a':
        print("Please select the destination folder for output files (e.g., where product.csv will be created).")
        dest_folder = select_folder_dialog("Select destination folder")
        if not dest_folder:
            print("No destination folder selected. Exiting.")
            exit(1)
        print("Please select the source folder to run exiftool on (your files to be scanned).")
        source_folder = select_folder_dialog("Select source folder")
        if not source_folder:
            print("No source folder selected. Exiting.")
            exit(1)

        run_exiftool_and_create_source2_csv(dest_folder, source_folder)
        source1_path = extract_and_rename_zip(dest_folder)
        source2_path = os.path.join(dest_folder, "source2.csv")
        output_path = os.path.join(dest_folder, "product.csv")
        error_path = os.path.join(dest_folder, "error.txt")
        member_of_existing_entity_id = input("Enter value for member_of_existing_entity_id: ").strip()
        source2_mapping = load_source2(source2_path)
        source1_to_product(source1_path, source2_mapping, output_path, member_of_existing_entity_id, error_path)
        print("product.csv generated successfully.")
        print("error.txt written for unmatched rows and blank fields.")

    elif mode == 'm':
        print("Please select the folder containing source1.csv and (optionally) source2.csv.")
        map_folder = select_folder_dialog("Select folder with source1.csv and (optionally) source2.csv")
        if not map_folder:
            print("No folder selected. Exiting.")
            exit(1)
        source1_path = os.path.join(map_folder, "source1.csv")
        source2_path = os.path.join(map_folder, "source2.csv")
        output_file_name = input("Enter desired output file name for product CSV (e.g. product, product_v2, etc.): ").strip()
        if not output_file_name:
            output_file_name = "product.csv"
        else:
            if not output_file_name.lower().endswith(".csv"):
                output_file_name = f"{output_file_name}.csv"
        output_path = os.path.join(map_folder, output_file_name)
        mapping_report_path = os.path.join(map_folder, "mapping_report.txt")
        if not os.path.exists(source1_path):
            print("source1.csv not found in selected folder.")
            exit(1)
        if not os.path.exists(source2_path):
            print("source2.csv not found in selected folder.")
            print("Do you want to create source2.csv using exiftool? (y/n)")
            answer = input().strip().lower()
            if answer == "y":
                print("Please select the source folder to run exiftool on (your image files to be scanned).")
                image_folder = select_folder_dialog("Select source folder for exiftool")
                if not image_folder:
                    print("No source folder selected. Exiting.")
                    exit(1)
                run_exiftool_and_create_source2_csv(map_folder, image_folder)
            else:
                print("Cannot proceed without source2.csv. Exiting.")
                exit(1)
        member_of_existing_entity_id = input("Enter value for member_of_existing_entity_id (default 10678): ").strip() or "10678"
        maps_mode_generate_product(
            source1_path,
            source2_path,
            output_path,
            mapping_report_path,
            member_of_existing_entity_id=member_of_existing_entity_id
        )
        print(f"{os.path.basename(output_path)} generated successfully.")

        # Write missing_metadata.txt report after product.csv is generated
        missing_metadata_report_path = os.path.join(map_folder, "missing_metadata.txt")
        write_missing_metadata_report(output_path, source2_path, missing_metadata_report_path)
        print(f"Missing metadata report written to {os.path.basename(missing_metadata_report_path)}.")

        # Cleanup prompt logic
        cleanup = input("Would you like to delete source2.csv, and source1_cleaned.csv from the folder? (y/n): ").strip().lower()
        if cleanup == "y":
            files_to_delete = [
                os.path.join(map_folder, "source2.csv"),
                os.path.join(map_folder, "source1_cleaned.csv"),
            ]
            for fp in files_to_delete:
                try:
                    if os.path.exists(fp):
                        os.remove(fp)
                        print(f"Deleted {fp}")
                except Exception as e:
                    print(f"Could not delete {fp}: {e}")
        else:
            print("Cleanup skipped.")

    else:
        print("Invalid selection. Exiting.")
