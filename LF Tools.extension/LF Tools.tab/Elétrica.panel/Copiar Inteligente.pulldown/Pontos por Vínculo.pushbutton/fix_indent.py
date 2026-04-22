import sys

with open('script.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if line.startswith('            for item in to_place:'):
        start_idx = i
        break

if start_idx != -1:
    for i in range(start_idx + 1, len(lines)):
        line = lines[i]
        # End loop when we hit the carimbar part which is at 12 spaces
        if line.startswith('            # Carimbar para handshake'):
            end_idx = i
            break

if start_idx != -1 and end_idx != -1:
    new_lines = lines[:start_idx]
    new_lines.append('            total_inst = sum(len(x["instances"]) for x in to_place)\n')
    new_lines.append('            processed_inst = 0\n')
    new_lines.append('            with forms.ProgressBar(title="Colocando Pontos...", cancellable=True) as pb:\n')
    
    for i in range(start_idx, end_idx):
        if 'for item in to_place:' in lines[i]:
            new_lines.append('                for item in to_place:\n')
            new_lines.append('                    if pb.cancelled: break\n')
        elif 'for inst in item["instances"]:' in lines[i]:
            new_lines.append('                    for inst in item["instances"]:\n')
            new_lines.append('                        if pb.cancelled: break\n')
            new_lines.append('                        processed_inst += 1\n')
            new_lines.append('                        pb.update_progress(processed_inst, total_inst)\n')
        else:
            if lines[i].strip():
                new_lines.append('    ' + lines[i])
            else:
                new_lines.append(lines[i])
                
    new_lines.extend(lines[end_idx:])
    
    with open('script.py', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print('Progress Bar added and block indented.')
else:
    print('Failed to find block boundaries. start:', start_idx, 'end:', end_idx)
