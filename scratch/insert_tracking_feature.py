import re

TRACKING_CSS = """
        /* ==========================================================================
           LIVE FLEET TRACKING LEAFLET & RADAR CUSTOM STYLES
           ========================================================================== */
        .leaflet-container {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
            background: #0b1120 !important;
        }
        .leaflet-popup-content-wrapper, .leaflet-popup-tip {
            background: #0f172a !important;
            color: #f8fafc !important;
            border: 1px solid #334155 !important;
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.7) !important;
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
            box-shadow: 0 0 16px rgba(56, 189, 248, 0.85);
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
            box-shadow: 0 0 14px rgba(16, 185, 129, 0.9);
            font-size: 15px;
            transition: transform 0.4s ease;
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
            box-shadow: 0 3px 10px rgba(0,0,0,0.6);
            color: #ffffff;
            cursor: pointer;
        }
        .drop-pin-completed {
            background: #22c55e;
            box-shadow: 0 0 10px rgba(34, 197, 94, 0.7);
        }
        .drop-pin-intransit {
            background: #0284c7;
            box-shadow: 0 0 14px rgba(56, 189, 248, 0.9);
            animation: pulse 1.5s infinite;
        }
        .drop-pin-pending {
            background: #f59e0b;
            box-shadow: 0 0 6px rgba(245, 158, 11, 0.5);
        }
        .drop-pin-failed {
            background: #ef4444;
            box-shadow: 0 0 8px rgba(239, 68, 68, 0.6);
        }
"""

TRACKING_HTML_SECTION = """
                <!-- ==========================================================================
                     TAB: LIVE FLEET TRACKING & GPS TELEMATICS
                     ========================================================================== -->
                <section id="tracking-tab" class="tab-pane">
                    <!-- Live Tracking Control Bar & KPI Badges -->
                    <div class="card" style="margin-bottom: 18px; padding: 16px 20px; background: rgba(15, 23, 42, 0.85); border: 1px solid rgba(56, 189, 248, 0.3); border-radius: var(--radius-lg);">
                        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px;">
                            <!-- Left: Title and Depot/Route Controls -->
                            <div style="display: flex; align-items: center; gap: 14px; flex-wrap: wrap;">
                                <div style="display: flex; align-items: center; gap: 10px;">
                                    <div style="width: 42px; height: 42px; border-radius: 10px; background: rgba(56, 189, 248, 0.15); border: 1px solid #38bdf8; display: flex; align-items: center; justify-content: center; font-size: 20px; color: #38bdf8;">
                                        <i class="fa-solid fa-satellite-dish fa-fade"></i>
                                    </div>
                                    <div>
                                        <h2 style="font-size: 1.25rem; font-weight: 800; color: #f8fafc; margin: 0; display: flex; align-items: center; gap: 8px;">
                                            Live GPS Fleet & Van Tracking
                                            <span class="badge" style="background: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid #22c55e; font-size: 0.72rem; padding: 2px 8px; border-radius: 12px;">
                                                <i class="fa-solid fa-circle" style="font-size: 6px; animation: pulse 1.5s infinite;"></i> REAL-TIME GPS
                                            </span>
                                        </h2>
                                        <p style="font-size: 0.82rem; color: #94a3b8; margin: 2px 0 0;">Interactive OpenStreetMap visualization of live vans, delivery drop sequences & cash collections</p>
                                    </div>
                                </div>

                                <!-- Depot Filter Selector -->
                                <div style="display: flex; align-items: center; gap: 8px; margin-left: 8px;">
                                    <label for="trackingDepotSelect" style="font-size: 0.85rem; font-weight: 700; color: #cbd5e1;"><i class="fa-solid fa-warehouse text-cyan"></i> Depot:</label>
                                    <select id="trackingDepotSelect" onchange="onTrackingDepotChange()" style="background: #0b1329; border: 1px solid #334155; color: #38bdf8; padding: 6px 12px; border-radius: 6px; font-weight: 700; font-size: 0.85rem; min-width: 220px;">
                                        <!-- Dynamically populated -->
                                    </select>
                                </div>

                                <!-- Route Filter Selector -->
                                <div style="display: flex; align-items: center; gap: 8px;">
                                    <label for="trackingRouteFilter" style="font-size: 0.85rem; font-weight: 700; color: #cbd5e1;"><i class="fa-solid fa-truck-fast text-green"></i> Route:</label>
                                    <select id="trackingRouteFilter" onchange="filterTrackingRoutesOnMap()" style="background: #0b1329; border: 1px solid #334155; color: #f8fafc; padding: 6px 12px; border-radius: 6px; font-weight: 600; font-size: 0.85rem; min-width: 190px;">
                                        <option value="ALL">All Active Fleet Vans</option>
                                    </select>
                                </div>
                            </div>

                            <!-- Right: Action Buttons (Refresh, Movement Simulator, Center Map) -->
                            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                                <button type="button" class="btn btn-sm btn-outline" id="btnRecenterTracking" onclick="recenterTrackingMap()" title="Fit Map to all vehicles & drops" style="padding: 7px 12px; font-size: 0.82rem; border-color: #475569;">
                                    <i class="fa-solid fa-crosshairs"></i> Center Bounds
                                </button>
                                <button type="button" class="btn btn-sm btn-outline" id="btnSimulateMovement" onclick="triggerTrackingSimulation()" title="Simulate GPS Movement of en-route vans" style="padding: 7px 14px; font-size: 0.82rem; border-color: #f59e0b; color: #f59e0b; background: rgba(245, 158, 11, 0.1);">
                                    <i class="fa-solid fa-play"></i> Simulate Movement
                                </button>
                                <button type="button" class="btn btn-sm btn-primary" id="btnRefreshTracking" onclick="loadLiveTrackingData(true)" style="padding: 7px 14px; font-size: 0.82rem;">
                                    <i class="fa-solid fa-rotate" id="trackingSpinIcon"></i> Sync Live GPS
                                </button>
                            </div>
                        </div>

                        <!-- KPI Mini Stats Row -->
                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-top: 14px; padding-top: 14px; border-top: 1px solid rgba(51, 65, 85, 0.6);">
                            <div style="background: rgba(11, 19, 41, 0.8); padding: 10px 14px; border-radius: 8px; border: 1px solid #1e293b; display: flex; align-items: center; gap: 12px;">
                                <div style="width: 36px; height: 36px; border-radius: 8px; background: rgba(56, 189, 248, 0.15); color: #38bdf8; display: flex; align-items: center; justify-content: center; font-size: 16px;">
                                    <i class="fa-solid fa-truck-moving"></i>
                                </div>
                                <div>
                                    <div style="font-size: 0.75rem; color: #94a3b8; font-weight: 600; text-transform: uppercase;">Active Vans En-Route</div>
                                    <div id="statTrackingVans" style="font-size: 1.25rem; font-weight: 800; color: #38bdf8;">--</div>
                                </div>
                            </div>

                            <div style="background: rgba(11, 19, 41, 0.8); padding: 10px 14px; border-radius: 8px; border: 1px solid #1e293b; display: flex; align-items: center; gap: 12px;">
                                <div style="width: 36px; height: 36px; border-radius: 8px; background: rgba(34, 197, 94, 0.15); color: #22c55e; display: flex; align-items: center; justify-content: center; font-size: 16px;">
                                    <i class="fa-solid fa-circle-check"></i>
                                </div>
                                <div>
                                    <div style="font-size: 0.75rem; color: #94a3b8; font-weight: 600; text-transform: uppercase;">Completed Drops</div>
                                    <div id="statTrackingCompleted" style="font-size: 1.25rem; font-weight: 800; color: #22c55e;">--</div>
                                </div>
                            </div>

                            <div style="background: rgba(11, 19, 41, 0.8); padding: 10px 14px; border-radius: 8px; border: 1px solid #1e293b; display: flex; align-items: center; gap: 12px;">
                                <div style="width: 36px; height: 36px; border-radius: 8px; background: rgba(245, 158, 11, 0.15); color: #f59e0b; display: flex; align-items: center; justify-content: center; font-size: 16px;">
                                    <i class="fa-solid fa-hourglass-half"></i>
                                </div>
                                <div>
                                    <div style="font-size: 0.75rem; color: #94a3b8; font-weight: 600; text-transform: uppercase;">Pending / In-Transit</div>
                                    <div id="statTrackingPending" style="font-size: 1.25rem; font-weight: 800; color: #f59e0b;">--</div>
                                </div>
                            </div>

                            <div style="background: rgba(11, 19, 41, 0.8); padding: 10px 14px; border-radius: 8px; border: 1px solid #1e293b; display: flex; align-items: center; gap: 12px;">
                                <div style="width: 36px; height: 36px; border-radius: 8px; background: rgba(168, 85, 247, 0.15); color: #a855f7; display: flex; align-items: center; justify-content: center; font-size: 16px;">
                                    <i class="fa-solid fa-bangladeshi-taka-sign"></i>
                                </div>
                                <div>
                                    <div style="font-size: 0.75rem; color: #94a3b8; font-weight: 600; text-transform: uppercase;">Total Delivery Value</div>
                                    <div id="statTrackingTotalAmount" style="font-size: 1.25rem; font-weight: 800; color: #c084fc;">--</div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Main Split Layout: Leaflet Interactive Map (Left) + Live Route Timeline Feed (Right) -->
                    <div style="display: grid; grid-template-columns: minmax(500px, 1.8fr) minmax(340px, 1.2fr); gap: 20px; align-items: start;">
                        <!-- Left: Leaflet OpenStreetMap Container -->
                        <div class="card" style="padding: 0; overflow: hidden; border: 1px solid #334155; border-radius: var(--radius-lg); background: #0f172a; box-shadow: 0 10px 30px rgba(0,0,0,0.5);">
                            <div style="padding: 12px 18px; background: rgba(15, 23, 42, 0.95); border-bottom: 1px solid #334155; display: flex; justify-content: space-between; align-items: center;">
                                <div style="display: flex; align-items: center; gap: 10px;">
                                    <i class="fa-solid fa-map text-cyan"></i>
                                    <span style="font-weight: 700; color: #f8fafc; font-size: 0.9rem;" id="trackingMapHeaderTitle">Interactive Fleet Radar</span>
                                </div>
                                <div style="display: flex; align-items: center; gap: 12px; font-size: 0.78rem;">
                                    <span style="display: inline-flex; align-items: center; gap: 4px; color: #94a3b8;"><span style="width: 10px; height: 10px; border-radius: 50%; background: #38bdf8; display: inline-block;"></span> Depot</span>
                                    <span style="display: inline-flex; align-items: center; gap: 4px; color: #94a3b8;"><span style="width: 10px; height: 10px; border-radius: 50%; background: #22c55e; display: inline-block;"></span> Delivered</span>
                                    <span style="display: inline-flex; align-items: center; gap: 4px; color: #94a3b8;"><span style="width: 10px; height: 10px; border-radius: 50%; background: #0284c7; display: inline-block;"></span> En-Route</span>
                                    <span style="display: inline-flex; align-items: center; gap: 4px; color: #94a3b8;"><span style="width: 10px; height: 10px; border-radius: 50%; background: #f59e0b; display: inline-block;"></span> Pending</span>
                                </div>
                            </div>
                            
                            <div id="liveTrackingMap" style="width: 100%; height: 640px; background: #0b1120;">
                                <!-- Leaflet Mount Point -->
                            </div>
                        </div>

                        <!-- Right: Live Route & Delivery Feed Sidebar -->
                        <div class="card" style="padding: 0; overflow: hidden; border: 1px solid #334155; border-radius: var(--radius-lg); background: rgba(15, 23, 42, 0.85); max-height: 702px; display: flex; flex-direction: column;">
                            <div style="padding: 14px 18px; background: rgba(15, 23, 42, 0.95); border-bottom: 1px solid #334155; display: flex; justify-content: space-between; align-items: center;">
                                <div style="display: flex; align-items: center; gap: 8px;">
                                    <i class="fa-solid fa-list-check text-green"></i>
                                    <span style="font-weight: 700; color: #f8fafc; font-size: 0.9rem;">Live Route Fleet & Drop Feeds</span>
                                </div>
                                <span id="trackingLastSyncTime" style="font-size: 0.75rem; color: #64748b; font-family: 'JetBrains Mono', monospace;">Updated just now</span>
                            </div>

                            <!-- Search box within drops -->
                            <div style="padding: 10px 14px; background: #0b1329; border-bottom: 1px solid #1e293b;">
                                <div style="position: relative;">
                                    <i class="fa-solid fa-magnifying-glass" style="position: absolute; left: 10px; top: 10px; color: #64748b; font-size: 0.82rem;"></i>
                                    <input type="text" id="trackingSearchOutlet" oninput="filterTrackingFeedList()" placeholder="Search outlet name, invoice or route..." style="width: 100%; padding: 6px 10px 6px 30px; background: #0f172a; border: 1px solid #334155; border-radius: 6px; color: #fff; font-size: 0.82rem;">
                                </div>
                            </div>

                            <!-- Scrollable list of active vehicles and their drop sequences -->
                            <div id="trackingRouteFeedContainer" style="overflow-y: auto; padding: 14px; flex: 1; display: flex; flex-direction: column; gap: 14px;">
                                <!-- Dynamically populated route cards -->
                                <div style="text-align: center; color: #94a3b8; padding: 40px 20px;">
                                    <i class="fa-solid fa-spinner fa-spin" style="font-size: 24px; color: #38bdf8; margin-bottom: 10px;"></i>
                                    <p>Loading live fleet telematics...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
"""

TRACKING_MODAL = """
        <!-- ==========================================================================
             MODAL: UPDATE DROP DELIVERY STATUS
             ========================================================================== -->
        <div id="updateDropStatusModal" style="display: none; position: fixed; inset: 0; background: rgba(11, 17, 32, 0.88); backdrop-filter: blur(8px); z-index: 100005; align-items: center; justify-content: center; padding: 20px;">
            <div style="background: var(--bg-card); border: 1px solid rgba(56, 189, 248, 0.4); border-radius: var(--radius-lg); width: 100%; max-width: 480px; padding: 24px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7);">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; border-bottom: 1px solid #334155; padding-bottom: 12px;">
                    <h3 style="font-size: 1.1rem; font-weight: 800; color: #fff; margin: 0; display: flex; align-items: center; gap: 8px;">
                        <i class="fa-solid fa-clipboard-check text-cyan"></i> Update Drop Status
                    </h3>
                    <button type="button" onclick="closeUpdateDropModal()" style="background: none; border: none; color: #94a3b8; font-size: 22px; cursor: pointer; line-height: 1;">&times;</button>
                </div>

                <form id="updateDropStatusForm" onsubmit="submitDropStatusUpdate(event)">
                    <input type="hidden" id="modalDropId">
                    <input type="hidden" id="modalRouteCode">
                    <input type="hidden" id="modalDepotId">

                    <div style="margin-bottom: 14px; background: #0b1329; padding: 12px; border-radius: 8px; border: 1px solid #1e293b;">
                        <div style="font-size: 0.8rem; color: #94a3b8;">Customer / Outlet:</div>
                        <div id="modalOutletName" style="font-size: 0.95rem; font-weight: 700; color: #38bdf8;">--</div>
                        <div id="modalOutletAddress" style="font-size: 0.78rem; color: #cbd5e1; margin-top: 2px;">--</div>
                    </div>

                    <div class="form-group" style="margin-bottom: 14px;">
                        <label for="modalDropStatusSelect" style="display: block; font-size: 0.85rem; font-weight: 700; margin-bottom: 6px; color: #cbd5e1;">Delivery Status *</label>
                        <select id="modalDropStatusSelect" required style="width: 100%; background: #0b1329; border: 1px solid #334155; color: #fff; padding: 8px 12px; border-radius: 6px; font-weight: 700;">
                            <option value="completed">Delivered (Completed)</option>
                            <option value="in_transit">In-Transit / En Route</option>
                            <option value="pending">Pending</option>
                            <option value="failed">Failed / Rescheduled</option>
                        </select>
                    </div>

                    <div class="form-group" style="margin-bottom: 14px;">
                        <label for="modalCashCollected" style="display: block; font-size: 0.85rem; font-weight: 700; margin-bottom: 6px; color: #cbd5e1;">Cash Collected (BDT ৳)</label>
                        <input type="number" id="modalCashCollected" step="0.01" min="0" placeholder="0.00" style="width: 100%; background: #0b1329; border: 1px solid #334155; color: #fff; padding: 8px 12px; border-radius: 6px;">
                    </div>

                    <div class="form-group" style="margin-bottom: 20px;">
                        <label for="modalProofNote" style="display: block; font-size: 0.85rem; font-weight: 700; margin-bottom: 6px; color: #cbd5e1;">Proof of Delivery / Gate Pass Note</label>
                        <input type="text" id="modalProofNote" placeholder="e.g. Received by Manager Hasan, signed delivery note" style="width: 100%; background: #0b1329; border: 1px solid #334155; color: #fff; padding: 8px 12px; border-radius: 6px;">
                    </div>

                    <div style="display: flex; justify-content: flex-end; gap: 10px;">
                        <button type="button" class="btn btn-outline" onclick="closeUpdateDropModal()">Cancel</button>
                        <button type="submit" class="btn btn-primary" style="font-weight: 700;">
                            <i class="fa-solid fa-check"></i> Save & Sync Telematics
                        </button>
                    </div>
                </form>
            </div>
        </div>
"""

TRACKING_JS_LOGIC = """
            // =========================================================================
            // LIVE FLEET TRACKING & GPS TELEMATICS CONTROLLER
            // =========================================================================
            let trackingMap = null;
            let trackingLayers = {
                depot: null,
                routes: null,
                drops: null,
                vehicles: null
            };
            let trackingData = null;
            let trackingAutoPollTimer = null;
            let dropMarkersRegistry = {};

            function initLiveTracking() {
                // 1. Populate tracking depot selector
                populateTrackingDepotDropdown();

                // 2. Initialize Leaflet Map if not created yet
                if (!trackingMap) {
                    const mapEl = document.getElementById('liveTrackingMap');
                    if (!mapEl) return;

                    trackingMap = L.map('liveTrackingMap', {
                        center: [23.7644, 90.3928],
                        zoom: 12,
                        zoomControl: true,
                        attributionControl: false
                    });

                    // Add dark/modern CartoDB Voyager tiles (fallback to OSM)
                    L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
                        subdomains: 'abcd',
                        maxZoom: 19
                    }).addTo(trackingMap);

                    // Create LayerGroups for clean rendering
                    trackingLayers.routes = L.layerGroup().addTo(trackingMap);
                    trackingLayers.drops = L.layerGroup().addTo(trackingMap);
                    trackingLayers.vehicles = L.layerGroup().addTo(trackingMap);
                    trackingLayers.depot = L.layerGroup().addTo(trackingMap);
                }

                // Invalidate map size after DOM transition
                setTimeout(() => {
                    if (trackingMap) trackingMap.invalidateSize();
                }, 150);

                // 3. Load tracking data
                loadLiveTrackingData(true);

                // 4. Start auto-refresh timer (every 15s)
                stopTrackingAutoPoll();
                trackingAutoPollTimer = setInterval(() => {
                    const trackingTab = document.getElementById('tracking-tab');
                    if (trackingTab && trackingTab.classList.contains('active')) {
                        loadLiveTrackingData(false);
                    } else {
                        stopTrackingAutoPoll();
                    }
                }, 15000);
            }

            function stopTrackingAutoPoll() {
                if (trackingAutoPollTimer) {
                    clearInterval(trackingAutoPollTimer);
                    trackingAutoPollTimer = null;
                }
            }

            function populateTrackingDepotDropdown() {
                const depotSelect = document.getElementById('trackingDepotSelect');
                if (!depotSelect) return;

                const currentVal = depotSelect.value;
                depotSelect.innerHTML = '';

                let depotsList = (window.currentDepots && window.currentDepots.length > 0) 
                    ? window.currentDepots 
                    : (window.masterDashboardData && window.masterDashboardData.depots ? window.masterDashboardData.depots : []);

                if (depotsList.length === 0) {
                    // Fallback list of default depots
                    depotsList = [
                        { id: 9, name: "Tejgaon Head Office & Central Depot", category: "Frozen/Ck" },
                        { id: 1, name: "Ashulia Factory - Frozen", category: "Frozen Food" },
                        { id: 2, name: "Gazipur Depot - Chicken", category: "Commercial Broiler" },
                        { id: 3, name: "Chittagong Central Depot", category: "Frozen & Layer" },
                        { id: 4, name: "Sylhet Regional Depot", category: "Poultry Feed & Chicks" }
                    ];
                }

                // If user is Depot Incharge, restrict
                const isDepotUser = currentUser && currentUser.role === 'incharge' && currentUser.depot_id;

                depotsList.forEach(d => {
                    if (isDepotUser && d.id !== currentUser.depot_id) return;
                    const opt = document.createElement('option');
                    opt.value = d.id;
                    opt.textContent = `${d.id}. ${d.name} (${d.category || 'General'})`;
                    depotSelect.appendChild(opt);
                });

                if (isDepotUser) {
                    depotSelect.value = currentUser.depot_id;
                    depotSelect.disabled = true;
                } else if (currentVal && depotSelect.querySelector(`option[value="${currentVal}"]`)) {
                    depotSelect.value = currentVal;
                } else if (depotsList.length > 0) {
                    depotSelect.value = depotsList[0].id;
                }
            }

            function onTrackingDepotChange() {
                loadLiveTrackingData(true);
            }

            async function loadLiveTrackingData(shouldFitBounds = false) {
                const depotSelect = document.getElementById('trackingDepotSelect');
                const depotId = depotSelect ? (depotSelect.value || 9) : 9;
                const spinIcon = document.getElementById('trackingSpinIcon');
                if (spinIcon) spinIcon.classList.add('fa-spin');

                try {
                    const res = await fetch(`/api/tracking/depot-routes?depot_id=${depotId}`);
                    const data = await res.json();

                    if (!data.success) {
                        showToast(data.message || 'Failed to load tracking data', 'error');
                        return;
                    }

                    trackingData = data;
                    renderTrackingMapData(data, shouldFitBounds);
                    renderTrackingFeedSidebar(data);
                    updateTrackingKPIs(data);

                    const syncTimeEl = document.getElementById('trackingLastSyncTime');
                    if (syncTimeEl) {
                        const now = new Date();
                        syncTimeEl.textContent = `Sync: ${now.toLocaleTimeString()}`;
                    }
                } catch (err) {
                    console.error('Error loading live tracking telematics:', err);
                } finally {
                    if (spinIcon) spinIcon.classList.remove('fa-spin');
                }
            }

            function updateTrackingKPIs(data) {
                const routes = data.routes || [];
                let totalDrops = 0;
                let completedDrops = 0;
                let totalValue = 0;

                routes.forEach(r => {
                    totalDrops += (r.total_drops || 0);
                    completedDrops += (r.completed_drops || 0);
                    (r.drop_points || []).forEach(d => {
                        totalValue += (parseFloat(d.total_amount) || 0);
                    });
                });

                const activeVansEl = document.getElementById('statTrackingVans');
                const completedEl = document.getElementById('statTrackingCompleted');
                const pendingEl = document.getElementById('statTrackingPending');
                const totalAmountEl = document.getElementById('statTrackingTotalAmount');

                if (activeVansEl) activeVansEl.textContent = `${routes.length} Vans`;
                if (completedEl) {
                    const pct = totalDrops > 0 ? Math.round((completedDrops / totalDrops) * 100) : 0;
                    completedEl.textContent = `${completedDrops} / ${totalDrops} (${pct}%)`;
                }
                if (pendingEl) pendingEl.textContent = `${Math.max(0, totalDrops - completedDrops)} Drops`;
                if (totalAmountEl) totalAmountEl.textContent = `৳ ${totalValue.toLocaleString('en-US', { maximumFractionDigits: 0 })}`;

                // Update Route Filter Dropdown
                const routeFilter = document.getElementById('trackingRouteFilter');
                if (routeFilter) {
                    const currentFilterVal = routeFilter.value;
                    routeFilter.innerHTML = '<option value="ALL">All Active Fleet Vans</option>';
                    routes.forEach(r => {
                        const opt = document.createElement('option');
                        opt.value = r.route_code;
                        opt.textContent = `${r.vehicle_no} - ${r.route_name}`;
                        routeFilter.appendChild(opt);
                    });
                    if (currentFilterVal && routeFilter.querySelector(`option[value="${currentFilterVal}"]`)) {
                        routeFilter.value = currentFilterVal;
                    }
                }
            }

            function renderTrackingMapData(data, shouldFitBounds = false) {
                if (!trackingMap) return;

                // Clear existing markers
                trackingLayers.depot.clearLayers();
                trackingLayers.routes.clearLayers();
                trackingLayers.drops.clearLayers();
                trackingLayers.vehicles.clearLayers();
                dropMarkersRegistry = {};

                const depotOrigin = data.depot_origin || { lat: 23.7644, lng: 90.3928, name: data.depot_name || "Central Depot" };
                const routes = data.routes || [];
                const allPoints = [[depotOrigin.lat, depotOrigin.lng]];

                // 1. Render Depot Marker
                const depotIcon = L.divIcon({
                    className: 'custom-depot-marker-wrapper',
                    html: `<div class="custom-depot-pin"><i class="fa-solid fa-warehouse"></i></div>`,
                    iconSize: [38, 38],
                    iconAnchor: [19, 19]
                });

                const depotMarker = L.marker([depotOrigin.lat, depotOrigin.lng], { icon: depotIcon }).addTo(trackingLayers.depot);
                depotMarker.bindPopup(`
                    <div style="font-weight: 800; color: #38bdf8; font-size: 0.95rem; margin-bottom: 4px;">
                        <i class="fa-solid fa-warehouse"></i> ${depotOrigin.name}
                    </div>
                    <div style="font-size: 0.8rem; color: #94a3b8;">
                        Central Dispatch Origin & Fleet Hub<br>
                        <strong>Active Routes:</strong> ${routes.length} Vans En-Route
                    </div>
                `);

                const routeColors = ['#38bdf8', '#10b981', '#f59e0b', '#a855f7', '#ec4899', '#06b6d4'];

                // 2. Render each route, drop points & vehicles
                routes.forEach((r, rIdx) => {
                    const color = routeColors[rIdx % routeColors.length];
                    const routeCoordinates = [[depotOrigin.lat, depotOrigin.lng]];

                    // Drops
                    (r.drop_points || []).forEach((drop, dIdx) => {
                        const dropCoord = [drop.lat, drop.lng];
                        routeCoordinates.push(dropCoord);
                        allPoints.push(dropCoord);

                        const st = (drop.status || 'pending').toLowerCase();
                        let pinClass = 'drop-pin-pending';
                        let stIcon = 'fa-hourglass-half';
                        let stText = 'Pending';
                        let stBadgeColor = '#f59e0b';

                        if (st === 'completed') {
                            pinClass = 'drop-pin-completed';
                            stIcon = 'fa-check';
                            stText = 'Delivered';
                            stBadgeColor = '#22c55e';
                        } else if (st === 'in_transit') {
                            pinClass = 'drop-pin-intransit';
                            stIcon = 'fa-truck-fast';
                            stText = 'En-Route';
                            stBadgeColor = '#0284c7';
                        } else if (st === 'failed') {
                            pinClass = 'drop-pin-failed';
                            stIcon = 'fa-xmark';
                            stText = 'Failed';
                            stBadgeColor = '#ef4444';
                        }

                        const dropIcon = L.divIcon({
                            className: 'custom-drop-marker-wrapper',
                            html: `<div class="custom-drop-pin ${pinClass}">${drop.sequence_order || (dIdx + 1)}</div>`,
                            iconSize: [28, 28],
                            iconAnchor: [14, 14]
                        });

                        const dropMarker = L.marker(dropCoord, { icon: dropIcon }).addTo(trackingLayers.drops);
                        dropMarkersRegistry[`${r.route_code}_${drop.id}`] = dropMarker;

                        dropMarker.bindPopup(`
                            <div style="min-width: 220px;">
                                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                                    <span style="font-weight:800; font-size:0.75rem; background:rgba(56,189,248,0.15); color:#38bdf8; padding:2px 6px; border-radius:4px;">
                                        Stop #${drop.sequence_order || (dIdx + 1)}
                                    </span>
                                    <span style="font-size:0.75rem; font-weight:700; color:${stBadgeColor};">
                                        <i class="fa-solid ${stIcon}"></i> ${stText}
                                    </span>
                                </div>
                                <div style="font-weight:800; font-size:0.95rem; color:#f8fafc; margin-bottom:3px;">
                                    ${drop.outlet_name}
                                </div>
                                <div style="font-size:0.78rem; color:#94a3b8; margin-bottom:6px;">
                                    <i class="fa-solid fa-location-dot text-cyan"></i> ${drop.address || 'Commercial Cluster'}
                                </div>
                                <div style="background:#0b1329; padding:6px 8px; border-radius:6px; font-size:0.78rem; margin-bottom:8px;">
                                    <div><strong>Inv:</strong> ${drop.invoice_no || 'N/A'} | <strong>Pkts:</strong> ${drop.total_pkts}</div>
                                    <div><strong>Value:</strong> ৳ ${parseFloat(drop.total_amount || 0).toLocaleString()} (${drop.payment_type || 'Cash'})</div>
                                </div>
                                <button type="button" onclick="openUpdateDropModal(${data.depot_id}, '${r.route_code}', '${drop.id}', '${escapeQuotes(drop.outlet_name)}', '${escapeQuotes(drop.address)}', '${st}', ${drop.total_amount || 0})" 
                                    style="width:100%; background:linear-gradient(135deg, #0284c7, #0369a1); color:#fff; border:none; padding:5px 8px; border-radius:4px; font-size:0.78rem; font-weight:700; cursor:pointer;">
                                    <i class="fa-solid fa-pen-to-square"></i> Update Status
                                </button>
                            </div>
                        `);
                    });

                    // Polyline Path
                    L.polyline(routeCoordinates, {
                        color: color,
                        weight: 4,
                        opacity: 0.85,
                        dashArray: '6, 8',
                        lineJoin: 'round'
                    }).addTo(trackingLayers.routes);

                    // Vehicle Marker
                    const vPos = r.vehicle_position || { lat: depotOrigin.lat, lng: depotOrigin.lng, heading: 45, speed: 30 };
                    allPoints.push([vPos.lat, vPos.lng]);

                    const vanIcon = L.divIcon({
                        className: 'custom-van-marker-wrapper',
                        html: `<div class="custom-van-pin" style="transform: rotate(${vPos.heading || 0}deg);"><i class="fa-solid fa-truck"></i></div>`,
                        iconSize: [36, 36],
                        iconAnchor: [18, 18]
                    });

                    const vanMarker = L.marker([vPos.lat, vPos.lng], { icon: vanIcon }).addTo(trackingLayers.vehicles);
                    vanMarker.bindPopup(`
                        <div style="min-width: 210px;">
                            <div style="font-weight: 800; color: #10b981; font-size: 0.95rem; margin-bottom: 2px;">
                                <i class="fa-solid fa-truck"></i> ${r.vehicle_no}
                            </div>
                            <div style="font-size: 0.82rem; color: #cbd5e1; font-weight: 600;">
                                ${r.route_name || r.route_code}
                            </div>
                            <hr style="border: 0; border-top: 1px solid #334155; margin: 6px 0;">
                            <div style="font-size: 0.78rem; color: #94a3b8; line-height: 1.6;">
                                <div><strong>Driver:</strong> ${r.driver_name} (<a href="tel:${r.driver_mobile}" style="color:#38bdf8;">${r.driver_mobile}</a>)</div>
                                <div><strong>Speed:</strong> <span style="color:#22c55e; font-weight:700;">${vPos.speed || 0} km/h</span></div>
                                <div><strong>Progress:</strong> ${r.completed_drops || 0} / ${r.total_drops || 0} Drops Delivered</div>
                            </div>
                        </div>
                    `);
                });

                // Auto-fit bounds on initial load or manual refresh
                if (shouldFitBounds && allPoints.length > 0) {
                    try {
                        trackingMap.fitBounds(allPoints, { padding: [40, 40], maxZoom: 15 });
                    } catch (e) { }
                }
            }

            function renderTrackingFeedSidebar(data) {
                const container = document.getElementById('trackingRouteFeedContainer');
                if (!container) return;

                const routes = data.routes || [];
                if (routes.length === 0) {
                    container.innerHTML = `
                        <div style="text-align: center; color: #94a3b8; padding: 40px 20px;">
                            <i class="fa-solid fa-truck-ramp-box" style="font-size: 32px; color: #475569; margin-bottom: 12px;"></i>
                            <h4 style="color: #cbd5e1; margin: 0 0 6px;">No Active Deliveries Today</h4>
                            <p style="font-size: 0.82rem;">Use Route Planning tab to dispatch vans for this depot.</p>
                        </div>
                    `;
                    return;
                }

                let html = '';
                routes.forEach((r, rIdx) => {
                    const total = r.total_drops || (r.drop_points ? r.drop_points.length : 0);
                    const done = r.completed_drops || 0;
                    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
                    const vPos = r.vehicle_position || { speed: 0 };

                    html += `
                        <div class="route-feed-card" style="background: #0b1329; border: 1px solid #1e293b; border-radius: 8px; padding: 12px; transition: border-color 0.2s;" id="routeCard_${r.route_code}">
                            <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 8px;">
                                <div>
                                    <div style="font-weight: 800; color: #38bdf8; font-size: 0.95rem; display: flex; align-items: center; gap: 6px;">
                                        <i class="fa-solid fa-truck-moving"></i> ${r.vehicle_no}
                                    </div>
                                    <div style="font-size: 0.78rem; color: #cbd5e1; margin-top: 1px;">
                                        ${r.route_name || r.route_code}
                                    </div>
                                </div>
                                <span class="badge" style="background: rgba(34, 197, 94, 0.15); color: #22c55e; border: 1px solid #22c55e; font-size: 0.72rem; padding: 2px 6px; border-radius: 4px;">
                                    ${vPos.speed || 0} km/h
                                </span>
                            </div>

                            <div style="font-size: 0.78rem; color: #94a3b8; margin-bottom: 8px;">
                                <i class="fa-solid fa-user text-cyan"></i> ${r.driver_name} | <a href="tel:${r.driver_mobile}" style="color: #38bdf8;"><i class="fa-solid fa-phone"></i> ${r.driver_mobile}</a>
                            </div>

                            <!-- Progress Bar -->
                            <div style="margin-bottom: 10px;">
                                <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: #94a3b8; margin-bottom: 4px;">
                                    <span>Delivery Progress</span>
                                    <span style="font-weight: 700; color: #f8fafc;">${done}/${total} Drops (${pct}%)</span>
                                </div>
                                <div style="height: 6px; background: #1e293b; border-radius: 3px; overflow: hidden;">
                                    <div style="height: 100%; width: ${pct}%; background: linear-gradient(90deg, #38bdf8, #22c55e); border-radius: 3px;"></div>
                                </div>
                            </div>

                            <!-- Drop Timeline List -->
                            <div style="display: flex; flex-direction: column; gap: 6px; margin-top: 8px;">
                    `;

                    (r.drop_points || []).forEach((drop, dIdx) => {
                        const st = (drop.status || 'pending').toLowerCase();
                        let stBadge = '<span style="color:#f59e0b; font-size:0.72rem;"><i class="fa-solid fa-hourglass-half"></i> Pending</span>';
                        if (st === 'completed') {
                            stBadge = '<span style="color:#22c55e; font-size:0.72rem;"><i class="fa-solid fa-check"></i> Delivered</span>';
                        } else if (st === 'in_transit') {
                            stBadge = '<span style="color:#38bdf8; font-size:0.72rem;"><i class="fa-solid fa-truck-fast"></i> En Route</span>';
                        } else if (st === 'failed') {
                            stBadge = '<span style="color:#ef4444; font-size:0.72rem;"><i class="fa-solid fa-xmark"></i> Failed</span>';
                        }

                        html += `
                            <div class="drop-item-row" data-outlet="${(drop.outlet_name || '').toLowerCase()}" data-invoice="${(drop.invoice_no || '').toLowerCase()}" style="background: rgba(15, 23, 42, 0.7); border: 1px solid #1e293b; padding: 8px 10px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;" onclick="panToOutletOnMap(${drop.lat}, ${drop.lng}, '${r.route_code}', '${drop.id}')">
                                <div style="display: flex; align-items: center; gap: 8px;">
                                    <span style="width: 20px; height: 20px; border-radius: 50%; background: #1e293b; color: #fff; font-size: 0.7rem; font-weight: 700; display: flex; align-items: center; justify-content: center;">
                                        ${drop.sequence_order || (dIdx + 1)}
                                    </span>
                                    <div>
                                        <div style="font-weight: 700; color: #f8fafc; font-size: 0.82rem;">${drop.outlet_name}</div>
                                        <div style="font-size: 0.72rem; color: #64748b;">${drop.invoice_no || ''} • ৳${parseFloat(drop.total_amount || 0).toLocaleString()}</div>
                                    </div>
                                </div>
                                <div style="text-align: right;">
                                    ${stBadge}
                                </div>
                            </div>
                        `;
                    });

                    html += `
                            </div>
                        </div>
                    `;
                });

                container.innerHTML = html;
            }

            function filterTrackingFeedList() {
                const searchInp = document.getElementById('trackingSearchOutlet');
                const query = searchInp ? searchInp.value.trim().toLowerCase() : '';
                const rows = document.querySelectorAll('.drop-item-row');

                rows.forEach(r => {
                    const outlet = r.dataset.outlet || '';
                    const invoice = r.dataset.invoice || '';
                    if (!query || outlet.includes(query) || invoice.includes(query)) {
                        r.style.display = 'flex';
                    } else {
                        r.style.display = 'none';
                    }
                });
            }

            function filterTrackingRoutesOnMap() {
                const routeFilter = document.getElementById('trackingRouteFilter');
                const selectedCode = routeFilter ? routeFilter.value : 'ALL';
                if (!trackingData) return;

                if (selectedCode === 'ALL') {
                    renderTrackingMapData(trackingData, true);
                } else {
                    const filteredData = {
                        ...trackingData,
                        routes: (trackingData.routes || []).filter(r => r.route_code === selectedCode)
                    };
                    renderTrackingMapData(filteredData, true);
                }
            }

            function panToOutletOnMap(lat, lng, routeCode, dropId) {
                if (!trackingMap) return;
                trackingMap.setView([lat, lng], 16, { animate: true });
                const markerKey = `${routeCode}_${dropId}`;
                const marker = dropMarkersRegistry[markerKey];
                if (marker) {
                    marker.openPopup();
                }
            }

            function recenterTrackingMap() {
                if (trackingData) {
                    renderTrackingMapData(trackingData, true);
                }
            }

            async function triggerTrackingSimulation() {
                const depotSelect = document.getElementById('trackingDepotSelect');
                const depotId = depotSelect ? (depotSelect.value || 9) : 9;
                const btn = document.getElementById('btnSimulateMovement');
                if (btn) btn.disabled = true;

                try {
                    const res = await fetch('/api/tracking/simulate-movement', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ depot_id: depotId })
                    });
                    const resJson = await res.json();
                    if (resJson.success) {
                        showToast('Simulation step applied! GPS telemetry updated.', 'success');
                        await loadLiveTrackingData(false);
                    }
                } catch (e) {
                    showToast('Simulation failed: ' + e.message, 'error');
                } finally {
                    if (btn) btn.disabled = false;
                }
            }

            function openUpdateDropModal(depotId, routeCode, dropId, outletName, address, currentStatus, amount) {
                const modal = document.getElementById('updateDropStatusModal');
                if (!modal) return;

                document.getElementById('modalDepotId').value = depotId;
                document.getElementById('modalRouteCode').value = routeCode;
                document.getElementById('modalDropId').value = dropId;
                document.getElementById('modalOutletName').textContent = outletName;
                document.getElementById('modalOutletAddress').textContent = address || 'Central Corridor Cluster';
                document.getElementById('modalDropStatusSelect').value = currentStatus || 'completed';
                document.getElementById('modalCashCollected').value = amount || '';
                document.getElementById('modalProofNote').value = '';

                modal.style.display = 'flex';
            }

            function closeUpdateDropModal() {
                const modal = document.getElementById('updateDropStatusModal');
                if (modal) modal.style.display = 'none';
            }

            async function submitDropStatusUpdate(e) {
                e.preventDefault();
                const depotId = document.getElementById('modalDepotId').value;
                const routeCode = document.getElementById('modalRouteCode').value;
                const dropId = document.getElementById('modalDropId').value;
                const status = document.getElementById('modalDropStatusSelect').value;
                const cash = parseFloat(document.getElementById('modalCashCollected').value) || 0;
                const note = document.getElementById('modalProofNote').value;

                try {
                    const res = await fetch('/api/tracking/update-drop-status', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            depot_id: depotId,
                            route_code: routeCode,
                            drop_id: dropId,
                            status: status,
                            cash_collected: cash,
                            proof_note: note
                        })
                    });
                    const data = await res.json();
                    if (data.success) {
                        showToast(data.message || 'Status updated successfully', 'success');
                        closeUpdateDropModal();
                        await loadLiveTrackingData(false);
                    } else {
                        showToast(data.message || 'Failed to update drop status', 'error');
                    }
                } catch (err) {
                    showToast('Error updating status: ' + err.message, 'error');
                }
            }

            function escapeQuotes(str) {
                if (!str) return '';
                return String(str).replace(/'/g, "\\'").replace(/"/g, '&quot;');
            }
"""

def update_html_file(file_path):
    print("Processing", file_path, "...")
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 1. Add CSS if not present
    if '.custom-depot-pin' not in content:
        # insert before </style>
        content = content.replace('</style>', TRACKING_CSS + '\n    </style>', 1)
        print(" -> CSS injected")

    # 2. Add HTML Section if not present
    if 'id="tracking-tab"' not in content:
        # Insert before <!-- TAB 4: SOP & ANTI-THEFT PROTOCOLS --> or <section id="sop-tab"
        if '<section id="sop-tab"' in content:
            content = content.replace('<section id="sop-tab"', TRACKING_HTML_SECTION + '\n                <section id="sop-tab"', 1)
            print(" -> HTML section injected before sop-tab")
        else:
            print(" -> WARNING: could not find insertion point for tracking-tab section")

    # 3. Add Modal before </body> if not present
    if 'id="updateDropStatusModal"' not in content:
        content = content.replace('</body>', TRACKING_MODAL + '\n</body>', 1)
        print(" -> Modal injected before </body>")

    # 4. Update initTabs() to call initLiveTracking() and stopTrackingAutoPoll()
    # Check if initLiveTracking is already in initTabs
    if 'initLiveTracking()' not in content:
        old_init_tabs_case = "else if (target === 'plan-tab') {\n                                initRoutePlanner();\n                            }"
        new_init_tabs_case = """else if (target === 'plan-tab') {
                                initRoutePlanner();
                            } else if (target === 'tracking-tab') {
                                initLiveTracking();
                            }"""
        if old_init_tabs_case in content:
            content = content.replace(old_init_tabs_case, new_init_tabs_case, 1)
            print(" -> initTabs updated with tracking-tab handler")
        else:
            # Try regex replace
            pattern = r"(else if\s*\(\s*target === ['\"]plan-tab['\"]\s*\)\s*\{\s*initRoutePlanner\(\);\s*\})"
            if re.search(pattern, content):
                content = re.sub(pattern, r"\1 else if (target === 'tracking-tab') { initLiveTracking(); }", content)
                print(" -> initTabs updated via regex")

    # Also add stopTrackingAutoPoll() inside stopAutoRefresh()
    if 'stopTrackingAutoPoll()' not in content:
        content = content.replace('function stopAutoRefresh() {', 'function stopAutoRefresh() {\n                stopTrackingAutoPoll();', 1)
        print(" -> stopAutoRefresh updated with stopTrackingAutoPoll()")

    # 5. Add Tracking JS Logic before </script>
    if 'function initLiveTracking()' not in content:
        content = content.replace('</script>', TRACKING_JS_LOGIC + '\n</script>', 1)
        print(" -> Tracking JS logic injected before </script>")

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Done updating", file_path)

if __name__ == '__main__':
    update_html_file('templates/index.html')
    update_html_file('../Paragon_Distribution_Live_Portal.html')
