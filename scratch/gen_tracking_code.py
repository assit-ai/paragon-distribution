# Python script to generate the HTML and JS components for Live Fleet Tracking

css_code = """
        /* ==========================================================================
           LIVE FLEET TRACKING LEAFLET CUSTOM STYLES
           ========================================================================== */
        .leaflet-container {
            font-family: 'Inter', sans-serif !important;
            background: #0b1120 !important;
        }
        .leaflet-popup-content-wrapper, .leaflet-popup-tip {
            background: #0f172a !important;
            color: #f8fafc !important;
            border: 1px solid #334155 !important;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.6) !important;
            border-radius: 8px !important;
        }
        .leaflet-popup-content {
            margin: 12px 14px !important;
            line-height: 1.5 !important;
            font-size: 0.85rem !important;
        }
        .custom-depot-pin {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 38px;
            height: 38px;
            background: linear-gradient(135deg, #0284c7, #0369a1);
            color: #ffffff;
            border-radius: 50%;
            border: 3px solid #38bdf8;
            box-shadow: 0 0 15px rgba(56, 189, 248, 0.8);
            font-size: 16px;
        }
        .custom-van-pin {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, #10b981, #059669);
            color: #ffffff;
            border-radius: 50%;
            border: 2px solid #ffffff;
            box-shadow: 0 0 12px rgba(16, 185, 129, 0.8);
            font-size: 15px;
            transition: transform 0.3s ease;
        }
        .custom-drop-pin {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            font-weight: 800;
            font-size: 12px;
            border: 2px solid #ffffff;
            box-shadow: 0 2px 8px rgba(0,0,0,0.5);
            color: #ffffff;
        }
        .drop-pin-completed {
            background: #22c55e;
            box-shadow: 0 0 8px #22c55e;
        }
        .drop-pin-intransit {
            background: #0284c7;
            box-shadow: 0 0 10px #38bdf8;
            animation: pulse 1.5s infinite;
        }
        .drop-pin-pending {
            background: #f59e0b;
        }
        .drop-pin-failed {
            background: #ef4444;
        }
"""

print("CSS definition length:", len(css_code))
