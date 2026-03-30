const { createApp, ref, computed, onMounted, reactive } = Vue;

createApp({
    setup() {
        const currentTab = ref('upload-bom');

        // =============================================
        // BOM Viewer State
        // =============================================
        const viewData = ref({ level0: [], level1: [], level2: [], level3: [], instruction_summary: [] });
        const isViewLoading = ref(true);
        const selectedViewLot = ref(null);
        const openViewItems = ref(new Set());
        const showSemiModal = ref(false);
        const showRawModal = ref(false);

        // =============================================
        // CSV Upload State
        // =============================================
        const csvRows = ref([]);
        const csvFileName = ref('');
        const isDragOver = ref(false);
        const csvInput = ref(null);
        const csvLevel0 = reactive({
            productName: '',
            productCode: '',
            lotNo: '',
            targetQty: '',
            version: '',
            mfgDate: new Date().toISOString().slice(0, 10),
            requestTeam: '',
            purpose: ''
        });

        const isCsvLevel0Valid = computed(() => {
            return csvLevel0.productName && csvLevel0.productCode && csvLevel0.lotNo &&
                   csvLevel0.targetQty && csvLevel0.version && csvLevel0.mfgDate &&
                   csvLevel0.requestTeam && csvLevel0.purpose;
        });

        // =============================================
        // Lot 계층 구조 적용 로직
        // =============================================
        const applyLotHierarchy = () => {
            const rows = csvRows.value;
            if (!rows || rows.length === 0) return;

            const getVal = (row, keys) => {
                for (const k of keys) {
                    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
                        return String(row[k]).trim();
                    }
                }
                return '';
            };

            const colMaps = {
                level: ['Level', '레벨', 'lvl'],
                parentLot: ['상위 Lot', '상위Lot', '상위 LOT', '상위LOT', '상위Lot ', 'Parent Lot', '상급Lot'],
                parentConn: ['상위 연결', '상위연결', '상위 연결(Code)', '상위연결(Code)', '상위 연결 ', 'Parent', '상위 연결(명칭)'],
                allocLot: ['할당 Lot', '할당Lot', '할당 LOT', '할당LOT', 'Lot No.', 'LotNo', '할당 Lot ', '할당 LOT NO'],
                codeNo: ['Code No.', 'CodeNo', '코드번호', '코드 No.', '품목코드', 'Code No. ', 'Code'],
                name: ['명칭 / 구성품', '명칭/구성품', '원재료명', '구성품 명칭', '품목명', 'Description', '명칭']
            };

            // Process Level 1 -> 2 -> 3
            [1, 2, 3].forEach(currentLvl => {
                rows.forEach(row => {
                    const l = parseInt(getVal(row, colMaps.level) || '0');
                    if (l !== currentLvl) return;

                    if (l === 1) {
                        row['상위 Lot'] = csvLevel0.lotNo;
                    } else {
                        let pStr = getVal(row, colMaps.parentConn).toLowerCase();
                        if (!pStr) pStr = getVal(row, colMaps.parentLot).toLowerCase();
                        
                        if (!pStr) return;

                        const parentLvl = l - 1;
                        const parent = rows.find(p => {
                            const pl = parseInt(getVal(p, colMaps.level) || '0');
                            if (pl !== parentLvl) return false;
                            
                            const pCode = getVal(p, colMaps.codeNo).toLowerCase();
                            const pName = getVal(p, colMaps.name).toLowerCase();
                            
                            return pCode === pStr || pName === pStr || 
                                   (pCode && pStr.includes(pCode)) || (pName && pStr.includes(pName)) ||
                                   (pCode && pCode.includes(pStr)) || (pName && pName.includes(pStr));
                        });

                        if (parent) {
                            let pAlloc = getVal(parent, colMaps.allocLot);
                            if (pAlloc) {
                                row['상위 Lot'] = pAlloc;
                            }
                        }
                    }
                });
            });

            // Special Rule for NC Item BCM005: Copy from PC
            rows.forEach(row => {
                const l = parseInt(getVal(row, colMaps.level) || '0');
                const itemCode = getVal(row, colMaps.codeNo);
                const parentLot = getVal(row, colMaps.parentLot);

                if (l === 2 && itemCode === 'BCM005' && parentLot.toUpperCase().includes('NC')) {
                    const pcRow = rows.find(p => 
                        parseInt(getVal(p, colMaps.level) || '0') === 2 && 
                        getVal(p, colMaps.codeNo) === 'BCM005' && 
                        getVal(p, colMaps.parentLot).toUpperCase().includes('PC')
                    );

                    if (pcRow) {
                        row['할당 Lot'] = getVal(pcRow, colMaps.allocLot);
                        row['유효기간'] = pcRow['유효기간'] || '';
                        row['필요 수량'] = 0;
                        row['할당수량'] = 0;
                        row._isNCSpecial = true;
                    }
                }
            });
            
            csvRows.value = JSON.parse(JSON.stringify(rows));
            alert('상위 Lot 계층 구조 및 NC 특수 로직(BCM005)이 적용되었습니다.');
        };

        // =============================================
        // CSV 파싱 유틸리티
        // =============================================
        function parseCsvLine(line) {
            const result = [];
            let current = '';
            let inQuotes = false;
            for (let i = 0; i < line.length; i++) {
                const ch = line[i];
                if (ch === '"') {
                    if (!inQuotes) { inQuotes = true; }
                    else if (line[i + 1] === '"') { current += '"'; i++; }
                    else { inQuotes = false; }
                } else if (ch === ',' && !inQuotes) {
                    result.push(current); current = '';
                } else { current += ch; }
            }
            result.push(current);
            return result;
        }

        function cleanVal(v) {
            return String(v || '').replace(/^="(.*)"$/, '$1').replace(/^=(.+)$/, '$1').trim();
        }

        function parseCsvText(text) {
            const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim());
            if (lines.length < 2) return [];
            const headers = parseCsvLine(lines[0]).map(h => cleanVal(h));
            return lines.slice(1).map(line => {
                const vals = parseCsvLine(line);
                const row = {};
                headers.forEach((h, i) => { row[h] = cleanVal(vals[i] || ''); });
                return row;
            }).filter(r => headers.some(h => r[h]));
        }

        // =============================================
        // CSV Computed Properties
        // =============================================
        const csvFiltered = computed(() => csvRows.value);
        const csvByLevel = computed(() => {
            const result = {};
            csvFiltered.value.forEach(r => {
                const lvl = parseInt(r['Level']);
                if (!isNaN(lvl)) { if (!result[lvl]) result[lvl] = []; result[lvl].push(r); }
            });
            return result;
        });

        const csvByLevelGrouped = computed(() => {
            const result = {};
            Object.keys(csvByLevel.value).forEach(lvl => {
                const rows = csvByLevel.value[lvl];
                const groups = [];
                let i = 0;
                
                // 1차 그룹화: 상위 Lot + Code + 명칭
                while (i < rows.length) {
                    const row = rows[i];
                    const key = `${row['상위 Lot']}|${row['Code No.']}|${row['명칭 / 구성품']}`;
                    const group = [row];
                    let j = i + 1;
                    while (j < rows.length) {
                        const nk = `${rows[j]['상위 Lot']}|${rows[j]['Code No.']}|${rows[j]['명칭 / 구성품']}`;
                        if (nk === key) { group.push(rows[j]); j++; } else break;
                    }
                    groups.push({ rows: group, span: group.length, lot: row['상위 Lot'] });
                    i = j;
                }

                // 2차 처리: 동일한 '상위 Lot' 연속 그룹들의 rowspan(lotSpan) 계산
                let gIdx = 0;
                while (gIdx < groups.length) {
                    const currentLot = groups[gIdx].lot;
                    let totalRowsForLot = groups[gIdx].span;
                    let nextGIdx = gIdx + 1;
                    
                    while (nextGIdx < groups.length && groups[nextGIdx].lot === currentLot) {
                        totalRowsForLot += groups[nextGIdx].span;
                        nextGIdx++;
                    }
                    
                    groups[gIdx].isFirstInLot = true;
                    groups[gIdx].lotSpan = totalRowsForLot;
                    
                    for (let k = gIdx + 1; k < nextGIdx; k++) {
                        groups[k].isFirstInLot = false;
                    }
                    gIdx = nextGIdx;
                }
                
                result[lvl] = groups;
            });
            return result;
        });

        const isExpiryNear = (dateStr) => {
            if (!dateStr) return false;
            const d = new Date(dateStr);
            if (isNaN(d)) return false;
            const diffDays = (d - new Date()) / (1000 * 60 * 60 * 24);
            return diffDays > 0 && diffDays < 180;
        };

        // =============================================
        // CSV 파일 로드 / 다운로드
        // =============================================
        const loadCsvFile = (file) => {
            csvFileName.value = file.name;
            
            const filename = file.name;
            const codeNo = filename.substring(0, 5);
            const parts = filename.split('_');
            let targetQty = '';
            if (parts.length >= 3) {
                targetQty = parts[2];
            }
            
            fetch(`/api/item_master/${codeNo}`)
                .then(res => res.json())
                .then(data => {
                    if (data && !data.error) {
                        csvLevel0.productCode = data.code_no || codeNo;
                        csvLevel0.productName = data.description || '';
                        csvLevel0.targetQty = targetQty;
                    } else {
                        csvLevel0.productCode = codeNo;
                        csvLevel0.targetQty = targetQty;
                    }
                })
                .catch(e => console.error('item_master fetch error:', e));

            const tryParse = (encoding) => new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(parseCsvText(e.target.result));
                reader.readAsText(file, encoding);
            });
            tryParse('UTF-8').then(rows => {
                if (rows.length > 0 && rows[0]['상태']) {
                    csvRows.value = rows;
                } else {
                    tryParse('EUC-KR').then(rows2 => { csvRows.value = rows2; });
                }
            });
        };

        const triggerCsvInput = () => { if (csvInput.value) csvInput.value.click(); };
        const handleCsvDrop = (e) => { isDragOver.value = false; const f = e.dataTransfer.files[0]; if (f) loadCsvFile(f); };
        const handleCsvFile = (e) => { const f = e.target.files[0]; if (f) loadCsvFile(f); };
        const resetCsv = () => { csvRows.value = []; csvFileName.value = ''; };

        const downloadCsvResult = () => {
            if (!csvFiltered.value.length || typeof XLSX === 'undefined') return;
            const headers = ['Level', '상위 Lot', 'Code No.', '명칭 / 구성품', '필요 수량', '단위', '할당 Lot', '유효기간', '할당수량'];
            
            // Level 0 row construction (Production Info)
            const level0Row = [
                '0', 
                '', 
                csvLevel0.productCode, 
                csvLevel0.productName, 
                csvLevel0.targetQty, 
                'EA', 
                csvLevel0.lotNo, 
                csvLevel0.mfgDate, 
                csvLevel0.targetQty
            ];

            const wsData = [
                headers, 
                level0Row,
                ...csvFiltered.value.map(r => headers.map(h => r[h] || ''))
            ];

            const ws = XLSX.utils.aoa_to_sheet(wsData);
            const colWidths = [10, 16, 12, 35, 10, 6, 22, 12, 10, 8];
            ws['!cols'] = colWidths.map(w => ({ wch: w }));
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'BOM_불출현황');
            const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            XLSX.writeFile(wb, `BOM_불출현황_${date}.xlsx`);
        };

        // =============================================
        // BOM Viewer API & Logic
        // =============================================
        const loadViewerData = async () => {
            isViewLoading.value = true;
            try {
                const res = await fetch('/api/bom-all');
                const data = await res.json();
                Object.keys(data).forEach(key => {
                    data[key] = data[key].map((item, idx) => {
                        const cleaned = {};
                        Object.entries(item).forEach(([k, v]) => {
                            cleaned[k] = String(v ?? '').replace(/\s00:00:00(\.\d+)?/g, '').trim();
                        });
                        return { ...cleaned, ID: `${key}_${idx}` };
                    });
                });
                viewData.value = data;
                if (data.level0.length > 0) selectViewLot(data.level0[0]['LOT NO.']);
            } finally {
                isViewLoading.value = false;
            }
        };

        const selectViewLot = (lot) => { selectedViewLot.value = lot; openViewItems.value = new Set(); };
        const currentL0 = computed(() => viewData.value.level0.find(i => i['LOT NO.'] === selectedViewLot.value));
        const displayL0Info = computed(() => {
            if (!currentL0.value) return {};
            const { Level: _, ID: __, ...rest } = currentL0.value;
            return rest;
        });

        const toggleOpen = (id) => {
            const s = new Set(openViewItems.value);
            s.has(id) ? s.delete(id) : s.add(id);
            openViewItems.value = s;
        };
        const isOpen = (id) => openViewItems.value.has(id);
        const splitLines = (str) => String(str || '').split('\n').filter(Boolean);

        const groupItems = (items, keyCols, valCols) => {
            const groups = new Map();
            items.forEach(item => {
                const key = keyCols.map(c => item[c]).join('||');
                if (!groups.has(key)) {
                    const g = { ...item };
                    valCols.forEach(c => { g[c] = String(item[c] || ''); });
                    groups.set(key, g);
                } else {
                    const g = groups.get(key);
                    valCols.forEach(c => { g[c] += '\n' + String(item[c] || ''); });
                }
            });
            return Array.from(groups.values());
        };

        const filteredL1 = computed(() => {
            const l1 = viewData.value.level1.filter(i => i['상위Lot'] === selectedViewLot.value);
            return groupItems(l1, ['코드번호', '구성품 명칭'], ['Lot No.', '제조일자', '유효기간', '포장시 요구량']);
        });

        const getL2SubItems = (l1LotStr) => {
            const lots = splitLines(l1LotStr);
            const l2 = viewData.value.level2.filter(i => lots.includes(i['상위Lot']));
            return groupItems(l2, ['상위Lot', '코드번호', '원재료명', '제조사'], ['Lot No.', '제조일자', '유효기간', '제조량']);
        };

        const getL3SubItems = (l2Items) => {
            const l2Lots = l2Items.flatMap(i => splitLines(i['Lot No.']));
            const l3 = viewData.value.level3.filter(i => l2Lots.includes(i['상위Lot']));
            return groupItems(l3, ['상위Lot', '코드번호', '원재료명', '제조사'], ['Lot No.', '제조일자', '유효기간', '제조량']);
        };

        const hasChildren = (l1LotStr) => {
            const lots = splitLines(l1LotStr);
            return viewData.value.level2.some(i => lots.includes(i['상위Lot']));
        };

        const filteredInstructions = computed(() => viewData.value.instruction_summary.filter(i => i['상위Lot'] === selectedViewLot.value));
        
        const aggregatedMaterials = computed(() => {
            const l1Items = viewData.value.level1.filter(i => i['상위Lot'] === selectedViewLot.value);
            const l1Lots = l1Items.map(i => i['Lot No.']).filter(Boolean);
            const l2All = viewData.value.level2.filter(i => l1Lots.includes(i['상위Lot']));
            const l2Lots = l2All.map(i => i['Lot No.']).filter(Boolean);
            const l3All = viewData.value.level3.filter(i => l2Lots.includes(i['상위Lot']));

            const map = new Map();
            const proc = (items) => {
                items.forEach(item => {
                    const code = item['코드번호'] || '';
                    if (!code) return;
                    if (!map.has(code)) map.set(code, { code, name: (item['원재료명'] || item['구성품 명칭'] || ''), total: 0, unit: item['단위'] || '' });
                    const q = parseFloat(String(item['제조량'] || item['포장시 요구량'] || '0').replace(/,/g, ''));
                    if (!isNaN(q)) map.get(code).total += q;
                });
            };
            proc(l3All); proc(l2All); proc(l1Items);
            const res = Array.from(map.values());
            res.forEach(r => { if (r.code === 'BCM005') r.total /= 2; r.total = Math.round(r.total * 1000) / 1000; });
            return res.sort((a,b) => a.code.localeCompare(b.code));
        });

        // =============================================
        // 반제품 Lot Modal Logic
        // =============================================
        const showSemiLotModal = ref(false);
        const semiLotList = ref([]);

        const parseDateInput = (val) => {
            if (!val) return null;
            const s = String(val).replace(/\D/g, '');
            if (s.length === 8) {
                return new Date(`${s.substring(0,4)}-${s.substring(4,6)}-${s.substring(6,8)}`);
            } else if (s.length === 6) {
                return new Date(`20${s.substring(0,2)}-${s.substring(2,4)}-${s.substring(4,6)}`);
            }
            const d = new Date(val);
            return isNaN(d) ? null : d;
        };

        const calcSemiLot = (mfgDate, division, docNo) => {
            if (!mfgDate || !division || !docNo) return '';
            try {
                const dateObj = parseDateInput(mfgDate);
                if (!dateObj || isNaN(dateObj)) return '';
                const yy = String(dateObj.getFullYear()).slice(-2);
                const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
                const dd = String(dateObj.getDate()).padStart(2, '0');
                const textDate = `${mm}${dd}${yy}`;

                const divM1 = String(division).substring(3, 5);
                const divM2 = String(division).substring(0, 2);

                const docM = String(docNo).substring(19, 21);
                const docR = String(docNo).slice(-2);
                
                return `${textDate}-${divM1}${divM2}-${docM}${docR}`;
            } catch(e) { return ''; }
        };

        const calcExpiryDate = (mfgDate) => {
            if (!mfgDate) return '';
            const d = parseDateInput(mfgDate);
            if (!d || isNaN(d)) return '';
            d.setMonth(d.getMonth() + 13);
            d.setDate(d.getDate() - 1);
            const yyyy = d.getFullYear();
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            return `${yyyy}-${mm}-${dd}`;
        };

        const openSemiLotModal = async () => {
            if (!csvLevel0.productCode) {
                alert("먼저 Level 0의 제품 Code No.를 확인해주세요.");
                return;
            }
            try {
                const res = await fetch(`/api/doc_master/${csvLevel0.productCode}`);
                const data = await res.json();
                if (data && data.length) {
                    const filteredData = data.filter(d => {
                        const div = String(d.division || '').toUpperCase();
                        return !div.startsWith('LA') && !div.startsWith('PI');
                    });
                    semiLotList.value = filteredData.map(d => ({
                        ...d, mfgDate: '', calcLot: '', expiryDate: ''
                    }));
                    showSemiLotModal.value = true;
                } else {
                    alert('해당 Code No.에 대한 반제품 제조지침서를 찾을 수 없습니다.');
                }
            } catch (e) {
                alert('데이터 로드 실패: ' + e);
            }
        };

        const onSemiMfgDateChange = (item) => {
            if (item.mfgDate) {
                item.calcLot = calcSemiLot(item.mfgDate, item.division, item.latest_doc_no);
                item.expiryDate = calcExpiryDate(item.mfgDate);
            } else {
                item.calcLot = '';
                item.expiryDate = '';
            }
        };

        const applySemiLots = () => {
            semiLotList.value.forEach(semi => {
                if (!semi.calcLot && !semi.expiryDate) return;
                
                csvRows.value.forEach(row => {
                    if (String(row['Code No.'] || '').trim().toUpperCase() === String(semi.division || '').trim().toUpperCase()) {
                        if (semi.calcLot) row['할당 Lot'] = semi.calcLot;
                        if (semi.expiryDate) row['유효기간'] = semi.expiryDate;
                        if (row['필요 수량']) row['할당수량'] = row['필요 수량'];
                    }
                });
            });
            alert('일괄 적용되었습니다.');
            showSemiLotModal.value = false;
        };

        // =============================================
        // Lifecycle
        // =============================================
        onMounted(() => {
            loadViewerData();
        });

        return {
            currentTab, loadViewerData,
            viewData, isViewLoading, selectedViewLot, currentL0, displayL0Info, selectViewLot,
            filteredL1, isOpen, toggleOpen, hasChildren, getL2SubItems, getL3SubItems, splitLines,
            showSemiModal, showRawModal, filteredInstructions, aggregatedMaterials,
            // CSV Upload Tab
            csvRows, csvFileName, isDragOver, csvInput, csvLevel0, csvFiltered, csvByLevel, csvByLevelGrouped,
            triggerCsvInput, handleCsvDrop, handleCsvFile, resetCsv, downloadCsvResult, isExpiryNear,
            showSemiLotModal, semiLotList, openSemiLotModal, onSemiMfgDateChange, applySemiLots,
            isCsvLevel0Valid, applyLotHierarchy
        };
    }
}).mount('#app');
