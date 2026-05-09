# ============================================================
# seed.py — Test Data Script
# Run this once to populate your database with 10 test towels
# Make sure your server is running before you run this!
# Run with: python seed.py
# ============================================================

import requests

# The base URL of your running server
BASE_URL = "https://web-production-51ae0.up.railway.app"

# Define 10 test towels with different types
# Each one needs a unique tag_id — in real life this would be the RFID chip number
towels = [
    {"tag_id": "TOWEL-001", "towel_type": "bath"},
    {"tag_id": "TOWEL-002", "towel_type": "bath"},
    {"tag_id": "TOWEL-003", "towel_type": "bath"},
    {"tag_id": "TOWEL-004", "towel_type": "hand"},
    {"tag_id": "TOWEL-005", "towel_type": "hand"},
    {"tag_id": "TOWEL-006", "towel_type": "hand"},
    {"tag_id": "TOWEL-007", "towel_type": "pool"},
    {"tag_id": "TOWEL-008", "towel_type": "pool"},
    {"tag_id": "TOWEL-009", "towel_type": "pool"},
    {"tag_id": "TOWEL-010", "towel_type": "face"},
]

# Loop through the list and register each one
for towel in towels:
    response = requests.post(
    f"{BASE_URL}/towels",
    json=towel,
    headers={"X-API-Key": "hotel-linen-2026-xK9mP"}
)

    if response.status_code == 200:
        print(f"✓ Registered {towel['tag_id']} ({towel['towel_type']})")
    elif response.status_code == 400:
        print(f"⚠ Skipped {towel['tag_id']} — already registered")
    else:
        print(f"✗ Failed {towel['tag_id']} — {response.text}")

print("\nDone! Hit /inventory to see your stock.")
