import re

def verify(filename):
    print(f"=== Verifying {filename} ===")
    with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    checks = {
        "Leaflet CSS": "leaflet.css" in content,
        "Leaflet JS": "leaflet.js" in content,
        "Nav Tab Button": 'data-tab="tracking-tab"' in content,
        "Tab Section (#tracking-tab)": 'id="tracking-tab"' in content,
        "Map Div (#liveTrackingMap)": 'id="liveTrackingMap"' in content,
        "Update Drop Modal": 'id="updateDropStatusModal"' in content,
        "initLiveTracking JS": "function initLiveTracking()" in content,
        "loadLiveTrackingData JS": "function loadLiveTrackingData(" in content,
        "renderTrackingMapData JS": "function renderTrackingMapData(" in content,
        "triggerTrackingSimulation JS": "function triggerTrackingSimulation()" in content,
        "Tab Click Handler Linked": "target === 'tracking-tab'" in content
    }

    for k, v in checks.items():
        print(f"  [{'PASS' if v else 'FAIL'}] {k}")

if __name__ == '__main__':
    verify('templates/index.html')
    verify('../Paragon_Distribution_Live_Portal.html')
