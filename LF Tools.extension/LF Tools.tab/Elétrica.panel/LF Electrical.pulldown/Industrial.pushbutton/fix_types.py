# coding: utf-8
import os

script_path = r"C:\Users\Luís Fernando\AppData\Roaming\pyRevit\Extensions\LF Tools.extension\LF Tools.tab\Elétrica.panel\LF Electrical.pulldown\Industrial.pushbutton\script.py"

with open(script_path, "r", encoding="utf-8") as f:
    content = f.read()

target = """    # Unlock read-only family parameters BEFORE starting the main transaction
    unique_types = set()
    for r in refs_list:
        host = doc.GetElement(r.ElementId)
        if hasattr(host, "GetTypeId"):
            elem_type = doc.GetElement(host.GetTypeId())
            if elem_type:
                unique_types.add(elem_type)
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        for sub_id, _ in valid_pairs:
            sub = doc.GetElement(sub_id)
            if hasattr(sub, "GetTypeId"):
                sub_type = doc.GetElement(sub.GetTypeId())
                if sub_type:
                    unique_types.add(sub_type)
    """

replacement = """    # Unlock read-only family parameters BEFORE starting the main transaction
    unique_types = set()
    for r in refs_list:
        host = doc.GetElement(r.ElementId)
        if hasattr(host, "GetTypeId"):
            elem_type = doc.GetElement(host.GetTypeId())
            if elem_type:
                unique_types.add(elem_type)
        parent = getattr(host, "SuperComponent", None)
        if parent and hasattr(parent, "GetTypeId"):
            parent_type = doc.GetElement(parent.GetTypeId())
            if parent_type:
                unique_types.add(parent_type)
        valid_pairs = get_valid_electrical_elements(host, (Domain.DomainElectrical,))
        for sub_id, _ in valid_pairs:
            sub = doc.GetElement(sub_id)
            if hasattr(sub, "GetTypeId"):
                sub_type = doc.GetElement(sub.GetTypeId())
                if sub_type:
                    unique_types.add(sub_type)
            parent_sub = getattr(sub, "SuperComponent", None)
            if parent_sub and hasattr(parent_sub, "GetTypeId"):
                parent_sub_type = doc.GetElement(parent_sub.GetTypeId())
                if parent_sub_type:
                    unique_types.add(parent_sub_type)
    """

if target in content:
    content = content.replace(target, replacement)
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(content)
    print("SUCCESS")
else:
    print("TARGET NOT FOUND")
    # let's try line by line
    with open(script_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    start = None
    end = None
    for i, line in enumerate(lines):
        if "unique_types = set()" in line:
            start = i - 1
        if "if unique_types:" in line and start is not None:
            end = i
            break
    if start and end:
        new_lines = lines[:start] + [replacement] + lines[end:]
        with open(script_path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)
        print("SUCCESS VIA LINES")
