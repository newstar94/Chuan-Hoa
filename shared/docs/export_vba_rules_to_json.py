import os
import re
import json

base_dir = r'd:\chuan-hoa-the-thuc-workspace'
vba_dir = os.path.join(base_dir, 'shared', 'vba_extracted')
rules_dir = os.path.join(base_dir, 'shared', 'rules')
dict_dir = os.path.join(base_dir, 'shared', 'dictionaries')

os.makedirs(rules_dir, exist_ok=True)
os.makedirs(dict_dir, exist_ok=True)

# 1. Parse RuleData.bas.bas for strings and arrays
rule_data_file = os.path.join(vba_dir, 'RuleData.bas.bas')
with open(rule_data_file, 'r', encoding='utf-8', errors='ignore') as f:
    vba_text = f.read()

print("Analyzing RuleData.bas.bas...")

# Extract key dictionaries and lists using regex
def extract_string_arrays(source):
    # Match patterns like: Result.Add "..." or Array("...", "...") or s(i) = "..."
    items = re.findall(r'"([^"\\]*(?:\\.[^"\\]*)*)"', source)
    return items

# Extract Typo Dictionary
typo_matches = re.findall(r'LoadRawTypoDictionary[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
typo_dict = {}
for block in typo_matches:
    pairs = re.findall(r'"([^"]+)"\s*,\s*"([^"]+)"', block)
    for wrong, right in pairs:
        typo_dict[wrong] = right

print(f"Extracted {len(typo_dict)} typo corrections")
with open(os.path.join(dict_dir, 'typo_dictionary.json'), 'w', encoding='utf-8') as f:
    json.dump(typo_dict, f, ensure_ascii=False, indent=2)

# Extract Iy Spellings
iy_matches = re.findall(r'LoadRawIyMapping[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
iy_dict = {}
for block in iy_matches:
    pairs = re.findall(r'"([^"]+)"\s*,\s*"([^"]+)"', block)
    for word_i, word_y in pairs:
        iy_dict[word_i] = word_y

print(f"Extracted {len(iy_dict)} i/y mappings")
with open(os.path.join(dict_dir, 'iy_dictionary.json'), 'w', encoding='utf-8') as f:
    json.dump(iy_dict, f, ensure_ascii=False, indent=2)

# Extract Non-Sentence-Ending Abbreviations (TS., PGS., v.v., ThS., ...)
non_ending_matches = re.findall(r'LoadRawNonSentenceEndingAbbreviations[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
non_ending_abbrs = []
for block in non_ending_matches:
    abbrs = re.findall(r'"([^"]+)"', block)
    non_ending_abbrs.extend(abbrs)
non_ending_abbrs = sorted(list(set(non_ending_abbrs)))

print(f"Extracted {len(non_ending_abbrs)} non-sentence-ending abbreviations")
with open(os.path.join(dict_dir, 'non_sentence_ending_abbreviations.json'), 'w', encoding='utf-8') as f:
    json.dump(non_ending_abbrs, f, ensure_ascii=False, indent=2)

# Extract DocType Abbreviations
doctype_matches = re.findall(r'LoadRawDocTypeAbbreviations[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
doctype_abbrs = {}
for block in doctype_matches:
    pairs = re.findall(r'"([^"]+)"\s*,\s*"([^"]+)"', block)
    for code, full in pairs:
        doctype_abbrs[code] = full

print(f"Extracted {len(doctype_abbrs)} doc type abbreviations")
with open(os.path.join(dict_dir, 'doctype_abbreviations.json'), 'w', encoding='utf-8') as f:
    json.dump(doctype_abbrs, f, ensure_ascii=False, indent=2)

# Extract Administrative Units & Place Names
admin_matches = re.findall(r'LoadRawAdministrativeUnitNames[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
admin_units = []
for block in admin_matches:
    units = re.findall(r'"([^"]+)"', block)
    admin_units.extend(units)
admin_units = sorted(list(set(admin_units)))

print(f"Extracted {len(admin_units)} administrative units")
with open(os.path.join(dict_dir, 'administrative_units.json'), 'w', encoding='utf-8') as f:
    json.dump(admin_units, f, ensure_ascii=False, indent=2)

# Extract Special Capitalizations
special_cap_matches = re.findall(r'LoadRawSpecialCapitalizations[^\n]*\n([\s\S]*?)End (?:Sub|Function)', vba_text)
special_caps = []
for block in special_cap_matches:
    caps = re.findall(r'"([^"]+)"', block)
    special_caps.extend(caps)
special_caps = sorted(list(set(special_caps)))

with open(os.path.join(dict_dir, 'special_capitalizations.json'), 'w', encoding='utf-8') as f:
    json.dump(special_caps, f, ensure_ascii=False, indent=2)

print("Dictionary extraction done!")
