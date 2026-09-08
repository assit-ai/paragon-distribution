import urllib.request
import urllib.parse
import json
import io
import zipfile
import sys

BASE_URL = "http://127.0.0.1:5000"

def log(msg, success=True):
    icon = "[PASS]" if success else "[FAIL]"
    print(f"{icon} {msg}")

def test_excel_export():
    print("\n--- 1. Testing Master Distribution Excel Report ---")
    url = f"{BASE_URL}/api/export/excel"
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as resp:
        status = resp.getcode()
        ctype = resp.headers.get('Content-Type', '')
        data = resp.read()
        
    assert status == 200, f"Expected 200, got {status}"
    assert 'spreadsheetml' in ctype or 'octet-stream' in ctype, f"Unexpected content-type: {ctype}"
    assert len(data) > 5000, f"Data size too small: {len(data)} bytes"
    
    zf = zipfile.ZipFile(io.BytesIO(data))
    namelist = zf.namelist()
    assert 'xl/workbook.xml' in namelist, "Not a valid Excel workbook"
    log(f"Excel report successfully generated! Size: {len(data):,} bytes. Sheets verified in archive.")

def test_admin_logins_and_clear_data():
    print("\n--- 2. Testing Role Logins and Admin Clear/Reseed Operations ---")
    
    # 2.1 Admin Login
    admin_login_data = json.dumps({"username": "admin", "password": "admin123"}).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/login", data=admin_login_data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as resp:
        admin_res = json.loads(resp.read().decode('utf-8'))
    assert admin_res.get('success') is True
    assert admin_res.get('user', {}).get('role') == 'admin'
    log(f"Admin login verified: role={admin_res['user']['role']}")

    # 2.2 Incharge Login
    inc_data = json.dumps({"username": "ashulia_frozen", "password": "depot123"}).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/login", data=inc_data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as resp:
        inc_res = json.loads(resp.read().decode('utf-8'))
    assert inc_res.get('success') is True
    assert inc_res.get('user', {}).get('role') == 'incharge'
    log(f"Incharge login verified: role={inc_res['user']['role']}, depot_id={inc_res['user']['depot_id']}")

    # 2.3 Delivery Man Login
    dm_data = json.dumps({"username": "rider_tejgaon", "password": "rider123"}).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/login", data=dm_data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as resp:
        dm_res = json.loads(resp.read().decode('utf-8'))
    assert dm_res.get('success') is True
    assert dm_res.get('user', {}).get('role') == 'delivery_man'
    log(f"Delivery Courier login verified: role={dm_res['user']['role']}, depot_id={dm_res['user']['depot_id']}")

    # 2.4 Clear All Demo Data
    clear_data = json.dumps({"role": "admin"}).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/admin/clear-all-demo-data", data=clear_data, headers={'Content-Type': 'application/json', 'X-Admin-Role': 'admin'})
    with urllib.request.urlopen(req) as resp:
        clear_res = json.loads(resp.read().decode('utf-8'))
    assert clear_res.get('success') is True
    log("Admin Clear Demo Data API returned success: " + clear_res.get('message', ''))

    # Verify summary is cleared
    with urllib.request.urlopen(f"{BASE_URL}/api/dashboard/summary?date=2026-09-08") as resp:
        summary_res = json.loads(resp.read().decode('utf-8'))
    compliance = summary_res.get('daily_entry_compliance', {})
    assert compliance.get('total_done') == 0, f"Expected 0 done, got {compliance.get('total_done')}"
    assert compliance.get('total_pending') == 17, f"Expected 17 pending, got {compliance.get('total_pending')}"
    log(f"Verified Dashboard shows 0 reports submitted ({compliance['total_pending']} pending across 17 depots).")

    # 2.5 Reseed Demo Data
    reseed_data = json.dumps({"role": "admin"}).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/admin/reseed-demo-data", data=reseed_data, headers={'Content-Type': 'application/json', 'X-Admin-Role': 'admin'})
    with urllib.request.urlopen(req) as resp:
        reseed_res = json.loads(resp.read().decode('utf-8'))
    assert reseed_res.get('success') is True
    log("Admin Reseed Demo Data API returned success: " + reseed_res.get('message', ''))

    # Verify summary is repopulated
    with urllib.request.urlopen(f"{BASE_URL}/api/dashboard/summary?date=2026-09-08") as resp:
        summary_res2 = json.loads(resp.read().decode('utf-8'))
    compliance2 = summary_res2.get('daily_entry_compliance', {})
    assert compliance2.get('total_done', 0) > 0, "Expected submitted reports after reseed"
    log(f"Verified Dashboard restored with {compliance2['total_done']} submitted reports.")

def test_delivery_courier_tracking_flow():
    print("\n--- 3. Testing Delivery Courier Drop Status & Real-Time Tracking Flow ---")
    
    with urllib.request.urlopen(f"{BASE_URL}/api/tracking/depot-routes?depot_id=1") as resp:
        trk_res = json.loads(resp.read().decode('utf-8'))
    assert trk_res.get('success') is True
    routes = trk_res.get('routes', [])
    assert len(routes) > 0, "Expected at least 1 route for depot 1"
    first_route = routes[0]
    r_code = first_route.get('route_code')
    drop_points = first_route.get('drop_points', [])
    assert len(drop_points) > 0, "Expected drop points in route"
    target_drop = drop_points[0]
    drop_id = target_drop['id']
    log(f"Fetched live route {r_code} with {len(drop_points)} drops. Target Drop ID: {drop_id}")

    # 3.2 Update Drop Status to 'completed' with cash collected
    update_payload = json.dumps({
        "depot_id": 1,
        "route_code": r_code,
        "drop_id": drop_id,
        "status": "completed",
        "cash_collected": 5250.0,
        "proof_note": "Handed to Store Manager Mr. Kamal, cash received."
    }).encode('utf-8')
    req = urllib.request.Request(f"{BASE_URL}/api/tracking/update-drop-status", data=update_payload, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as resp:
        upd_res = json.loads(resp.read().decode('utf-8'))
    assert upd_res.get('success') is True
    log(f"Drop update success: {upd_res.get('message')}")

    # 3.3 Verify updated status in tracking endpoint
    with urllib.request.urlopen(f"{BASE_URL}/api/tracking/depot-routes?depot_id=1") as resp:
        verify_trk = json.loads(resp.read().decode('utf-8'))
    v_routes = verify_trk.get('routes', [])
    v_first = next((r for r in v_routes if r['route_code'] == r_code), None)
    assert v_first is not None
    v_drop = next((d for d in v_first['drop_points'] if d['id'] == drop_id), None)
    assert v_drop is not None
    assert v_drop.get('status') == 'completed', f"Expected completed, got {v_drop.get('status')}"
    log(f"Verified live tracking feed reflects drop {drop_id} as status: {v_drop['status']}")

if __name__ == '__main__':
    try:
        test_excel_export()
        test_admin_logins_and_clear_data()
        test_delivery_courier_tracking_flow()
        print("\n=======================================================")
        print("ALL VERIFICATION TESTS PASSED SUCCESSFULLY!")
        print("=======================================================")
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        sys.exit(1)
