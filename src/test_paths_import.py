from src.common.io_paths import get_raw_root, get_out_root, get_log_root

print("RAW:", get_raw_root(must_exist=False))
print("OUT:", get_out_root())
print("LOG:", get_log_root())
