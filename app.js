const { createApp, ref, computed, onMounted, reactive, watch } = Vue;

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
        const viewDepth = ref(1); // 기본값: Level 1만 보기
        const showSemiModal = ref(false);
        const showRawModal = ref(false);

        const setViewDepth = (depth) => {
            viewDepth.value = depth;
            const newOpenItems = new Set();
            
            if (depth > 1) {
                // Level 2 이상일 경우, 하위 항목이 있는 Level 1 아코디언을 모두 엽니다.
                filteredL1.value.forEach(l1 => {
                    if (hasChildren(l1['Lot No.'])) {
                        newOpenItems.add(l1.ID);
                    }
                });
            }
            openViewItems.value = newOpenItems;
        };

        // =============================================
        // NAS 제조지침서 최신본 (R개정, 온디맨드)
        // =============================================
        const miLatestLoading = ref(false);
        const miLatestError = ref('');
        const miLatestResult = ref(null);
        const miManageFolder = ref('BCE01');

        const breadcrumbLabel = computed(() => {
            const labels = {
                viewer: 'BOM 조회',
                'upload-bom': '제조지시 실행',
                history: '제조지시 기록',
                'mi-manage': '제조지침서 관리'
            };
            return labels[currentTab.value] || '';
        });

        const selectedMiManageEntry = computed(() => {
            const data = miLatestResult.value;
            if (!data || !data.folders) return null;
            return data.folders[miManageFolder.value] ?? null;
        });

        const fetchManufacturingInstructionLatest = async () => {
            miLatestLoading.value = true;
            miLatestError.value = '';
            try {
                const res = await fetch('/api/manufacturing_instruction_latest');
                const data = await res.json();
                if (!res.ok) {
                    miLatestError.value = data.error || `HTTP ${res.status}`;
                    miLatestResult.value = null;
                    return;
                }
                miLatestResult.value = data;
            } catch (e) {
                miLatestError.value = String(e.message || e);
                miLatestResult.value = null;
            } finally {
                miLatestLoading.value = false;
            }
        };

        watch(currentTab, (tab) => {
            if (tab === 'mi-manage' && !miLatestLoading.value && !miLatestResult.value) {
                fetchManufacturingInstructionLatest();
            }
        });

        // =============================================
        // CSV Upload State
        // =============================================
        const csvRows = ref([]);
        const csvFileName = ref('');
        const isDragOver = ref(false);
        const csvInput = ref(null);
        const csvLevel0 = reactive({
            productName: '',      // 제품명 (item_master.description)
            modelName: '',        // 모델명 (A열: 코드번호)
            productInfo: '',      // 제품정보 (item_master.detailed_description)
            lotNo: '',            // LOT No. (수기)
            version: '',          // Version (item_master.version)
            targetQty: '',        // 제조수량 (파일명에서 파싱)
            mfgDate: new Date().toISOString().slice(0, 10), // 제조일자 (수기, 기본 오늘)
            expiryDate: '',       // 유효기간 (제조일자 +1년 -1일)
            requestTeam: '',      // 의뢰팀 (수기)
            purpose: ''           // 생산목적 (수기)
        });

        // 제조일자 변경 시 유효기간 자동 계산
        const calculateExpiry = () => {
            if (!csvLevel0.mfgDate) return;
            const d = new Date(csvLevel0.mfgDate);
            if (isNaN(d)) return;
            const day = d.getDate();
            d.setFullYear(d.getFullYear() + 1);
            if (d.getDate() !== day) d.setDate(0);
            d.setDate(d.getDate() - 1);
            csvLevel0.expiryDate = d.toISOString().slice(0, 10);
        };

        // 모델명 변경 시 마스터 정보 조회
        const fetchItemMasterDetail = async () => {
            if (!csvLevel0.modelName) return;
            try {
                const res = await fetch(`/api/item_master/${encodeURIComponent(csvLevel0.modelName)}`);
                const data = await res.json();
                if (data && !data.error) {
                    csvLevel0.productName = data.description || '';
                    csvLevel0.productInfo = data.detailed_description || '';
                    csvLevel0.version = data.version || '';
                }
            } catch (err) {
                console.error('Error fetching item master detail:', err);
            }
            // 기존 문서번호 조회 연동
            fetchDocMaster(csvLevel0.modelName);
        };

        // =============================================
        // Instruction Doc Master (완제품 패널용)
        // =============================================
        const docMasterList = ref([]);
        const piList = computed(() => {
            return docMasterList.value
                .filter(doc => doc.division && doc.division.toUpperCase().includes('PI'))
                .map(doc => {
                    return {
                        id: doc.id,
                        label: 'PI',
                        fullName: doc.division,
                        codeNo: doc.code_no,
                        latestDocNo: doc.latest_doc_no || '문서번호 없음'
                    };
                });
        });

        // 제품 코드 변경 시 문서 마스터 정보 조회
        const fetchDocMaster = async (codeNo) => {
            if (!codeNo) {
                docMasterList.value = [];
                return;
            }
            try {
                const res = await fetch(`/api/doc_master/${encodeURIComponent(codeNo)}`);
                const data = await res.json();
                docMasterList.value = Array.isArray(data) ? data : [];
            } catch (err) {
                console.error('Error fetching doc master:', err);
                docMasterList.value = [];
            }
        };

        // v-model 감시 (모델명 변경 시 마스터 정보 갱신)
        const onModelNameChange = () => {
            fetchItemMasterDetail();
        };

        const isCsvLevel0Valid = computed(() => {
            return csvLevel0.productName && csvLevel0.modelName && csvLevel0.lotNo &&
                   csvLevel0.targetQty && csvLevel0.version && csvLevel0.mfgDate &&
                   csvLevel0.requestTeam && csvLevel0.purpose && csvLevel0.productInfo &&
                   csvLevel0.expiryDate;
        });

        // =============================================
        // 제조지시 기록 (History) State
        // =============================================
        const historyLots = ref([]);
        const selectedHistoryLot = ref('');
        const historyDetail = ref(null);

        const historyDepth = ref(3); // 기본값: Level 3까지 모두 보기

        const setHistoryDepth = (depth) => {
            historyDepth.value = depth;
        };

        const loadHistoryLots = async () => {
            try {
                const res = await fetch('/api/instruction_lots');
                historyLots.value = await res.json();
            } catch (e) { console.error(e); }
        };

        const loadHistoryDetail = async () => {
            if (!selectedHistoryLot.value) return;
            try {
                const res = await fetch(`/api/instruction_detail/${selectedHistoryLot.value}`);
                historyDetail.value = await res.json();
            } catch (e) { alert('로드 실패: ' + e); }
        };

        const loadHistoryLotsRefresh = async () => {
            await loadHistoryLots();
        };

        // 데이터베이스에 저장
        const saveToDatabase = async () => {
            if (!confirm('현재 제조실시 내역을 데이터베이스에 저장하시겠습니까?')) return;
            
            // PI 항목을 instruction_summary 형식으로 변환하여 함께 저장
            const piSummaryItems = piList.value.map(pi => ({
                division: pi.fullName || 'PI',
                latest_doc_no: pi.latestDocNo,
                mfgDate: csvLevel0.mfgDate,
                calcLot: '', 
                expiryDate: ''
            }));

            const payload = {
                level0: csvLevel0,
                level1: csvRows.value.filter(r => parseInt(r['Level'] || r['level']) === 1),
                level2: csvRows.value.filter(r => parseInt(r['Level'] || r['level']) === 2),
                level3: csvRows.value.filter(r => parseInt(r['Level'] || r['level']) === 3),
                instruction_summary: [...piSummaryItems, ...semiLotList.value]
            };

            try {
                const res = await fetch('/api/save_instruction', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if (result.status === 'success') {
                    alert('성공적으로 저장되었습니다.');
                    loadHistoryLots(); 
                } else {
                    alert('저장 실패: ' + result.error);
                }
            } catch (err) {
                alert('서버 오류: ' + err);
            }
        };

        // 헬퍼: 행 그룹화 (기록 조회용)
        function groupRowsByLotAndCode(rows) {
            if (!rows || !rows.length) return [];
            const res = [];
            // DB 필드명 "상위Lot" 또는 "상위 Lot" 사용
            const lots = [...new Set(rows.map(r => r['상위Lot'] || r['상위 Lot']))];
            lots.forEach(lot => {
                const lotRows = rows.filter(r => (r['상위Lot'] || r['상위 Lot']) === lot);
                // DB 필드명 "코드번호" 사용
                const codes = [...new Set(lotRows.map(r => r['코드번호'] || r['Code No.']))];
                codes.forEach((code, cIdx) => {
                    const groupRows = lotRows.filter(r => (r['코드번호'] || r['Code No.']) === code);
                    
                    const totalQty = groupRows.reduce((sum, r) => {
                        const q = parseFloat(String(r['제조량'] || r['포장시 요구량'] || r['할당수량'] || r['필요 수량'] || '0').replace(/,/g, ''));
                        return sum + (isNaN(q) ? 0 : q);
                    }, 0);
                    const displayTotalQty = Math.round(totalQty * 1000) / 1000;

                    res.push({
                        isFirstInLot: cIdx === 0,
                        lotSpan: lotRows.length,
                        span: groupRows.length,
                        totalQty: displayTotalQty,
                        rows: groupRows
                    });
                });
            });
            return res;
        }

        const historyByLevelGrouped = computed(() => {
            if (!historyDetail.value) return {};
            const grouped = {};
            [1, 2, 3].forEach(lvl => {
                const rows = historyDetail.value[`level${lvl}`] || [];
                grouped[lvl] = groupRowsByLotAndCode(rows);
            });
            return grouped;
        });

        const historyPiList = computed(() => {
            if (!historyDetail.value) return [];
            const sum = historyDetail.value.instruction_summary || [];
            // DB 필드명 "약어", "제조지침서 No." 사용 (FI 또는 PI 포함 시 완제품으로 간주)
            return sum.filter(s => {
                const div = (s['약어'] || s['division'] || '').toUpperCase();
                return div.includes('PI') || div.includes('FI');
            }).map(s => ({ 
                label: 'PI', 
                latestDocNo: s['제조지침서 No.'] || s['latest_doc_no'] 
            }));
        });

        // =============================================
        // Packaging Instruction Modal State
        // =============================================
        const showPackagingModal = ref(false);
        const showProductManagementModal = ref(false); // 추가
        const productManagementPreview = ref({
            A7: '', I7: '', N7: '', T7: '', A9: '', I9: 0
        }); // 추가
        const packagingPreview = ref({
            E4: '', A7: '', J7: '', N7: 0, S7: '', Z7: '', AE7: '',
            items_mapped: []
        });

        const openPackagingInstruction = async (pi) => {
            let lotNo = '';
            if (currentTab.value === 'viewer') lotNo = selectedViewLot.value;
            else if (currentTab.value === 'upload-bom') lotNo = csvLevel0.lotNo;
            else if (currentTab.value === 'history') lotNo = selectedHistoryLot.value;

            if (!lotNo) {
                alert('대상 Lot No.를 먼저 선택하거나 생성해주세요.');
                return;
            }

            try {
                const res = await fetch(`/api/packaging_preview/${lotNo}`);
                const data = await res.json();
                if (data.error) {
                    alert('미리보기 데이터를 가져오는데 실패했습니다: ' + data.error);
                } else {
                    // 서버에서 받은 l1_items 데이터를 21-33 매핑 룰에 따라 가공
                    const mapping = {
                        'EMA015': { label: 'EMA015', row: 21 },
                        'EMA014': { label: 'EMA014', row: 22 },
                        'CR(01)': { label: 'CR(01)', row: 23 },
                        'PC(01)': { label: 'PC(01)', row: 24 },
                        'NC(01)': { label: 'NC(01)', row: 25 },
                        'DA(01)': { label: 'DA(01)', row: 26 },
                        'RD(01)': { label: 'RD(01)', row: 27 },
                        'WS(01)': { label: 'WS(01)', row: 28 },
                        'TM(01)': { label: 'TM(01)', row: 29 },
                        'SS(01)': { label: 'SS(01)', row: 30 },
                        'EMA013': { label: 'EMA013', row: 31 },
                        'PL(01)': { label: 'PL(01)', row: 32 },
                        'IFU': { label: 'IFU', row: 33 }
                    };

                    const l1_raw = data.EMA015_items || [];
                    const mapped = [];
                    const totalQty = data.N7 || 0;

                    l1_raw.forEach(item => {
                        const code = String(item['코드번호'] || '').toUpperCase();
                        for (const [key, val] of Object.entries(mapping)) {
                            if (code.includes(key)) {
                                if (val.row === 33) return; // 33행은 아래에서 별도 추가
                                mapped.push({
                                    ...val,
                                    lotNo: item['Lot No.'] || item['할당 Lot'],
                                    expiryDate: item['유효기간'],
                                    qty: item['포장시 요구량'] || item['할당수량'] || item['제조량']
                                });
                                break;
                            }
                        }
                    });

                    // 33행 명시적 추가 (L: 빈칸, S: 제조일자, X: 빈칸, AI: 총수량)
                    mapped.push({
                        label: 'IFU 등 기재',
                        row: 33,
                        lotNo: '-',
                        expiryDate: '-',
                        qty: totalQty
                    });

                    packagingPreview.value = { 
                        ...data,
                        items_mapped: mapped.sort((a, b) => a.row - b.row)
                    };
                    showPackagingModal.value = true;
                }
            } catch (err) {
                alert('서버 통신 오류: ' + err);
            }
        };

        const downloadPackagingFile = () => {
            const lotNo = packagingPreview.value.AE7;
            if (!lotNo) return;
            const link = document.createElement('a');
            link.href = `/api/packaging_download/${lotNo}`;
            link.setAttribute('download', ''); // 다운로드 강제
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        };

        const openProductManagement = async (pi) => {
            let lotNo = '';
            if (currentTab.value === 'viewer') lotNo = selectedViewLot.value;
            else if (currentTab.value === 'upload-bom') lotNo = csvLevel0.lotNo;
            else if (currentTab.value === 'history') lotNo = selectedHistoryLot.value;

            if (!lotNo) {
                alert('대상 Lot No.를 먼저 선택하거나 생성해주세요.');
                return;
            }

            try {
                const res = await fetch(`/api/product_management_preview/${lotNo}`);
                const data = await res.json();
                if (data.error) {
                    alert('미리보기 데이터를 가져오는데 실패했습니다: ' + data.error);
                } else {
                    productManagementPreview.value = data;
                    showProductManagementModal.value = true;
                }
            } catch (err) {
                alert('서버 통신 오류: ' + err);
            }
        };

        const downloadProductManagementFile = () => {
            const lotNo = productManagementPreview.value.N7;
            if (!lotNo) return;
            const link = document.createElement('a');
            link.href = `/api/product_management_download/${lotNo}`;
            link.setAttribute('download', ''); // 다운로드 강제
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        };

        const historySemiLotList = computed(() => {
            if (!historyDetail.value) return [];
            const sum = historyDetail.value.instruction_summary || [];
            return sum.filter(s => {
                const div = (s['약어'] || s['division'] || '').toUpperCase();
                // PI, FI, LA 제외 항목을 반제품으로 처리
                return div && !div.includes('PI') && !div.includes('FI') && !div.includes('LA');
            }).map(s => {
                const lot = s['Lot. No.'] || s['calcLot'] || s['calc_lot'] || '';
                return {
                    division: s['약어'] || s['division'],
                    calcLot: lot,
                    calc_lot: lot,
                    instructionNo: s['제조지침서 No.'] || s['latest_doc_no'] || '',
                    productionQty: s['생산량'] ?? '',
                    mfgDate: s['제조일자'] || '',
                    parentLot: s['상위Lot'] || ''
                };
            });
        });

        const showSemiMrModal = ref(false);
        const semiMrPreview = ref({
            division: '', instructionNo: '', lotNo: '', productionQty: '', mfgDate: '', l2: [], l3: [],
            /** PB·CB·WB: Level 3만 표시(Level 2 숨김). 그 외 반제품은 Level 2만(Level 3 미사용) */
            onlyLevel3Materials: false,
        });

        const semiMrUsesLevel3Only = (division) => {
            const d = String(division || '').toUpperCase().replace(/\s+/g, '');
            if (!d) return false;
            return d.startsWith('PB') || d.startsWith('CB') || d.startsWith('WB');
        };
        const showSemiMgmtModal = ref(false);
        const semiMgmtPreview = ref({
            A7: '', I7: '', N7: '', T7: '', A9: '', I9: '',
            division: '', instructionNo: '', lotNo: '', productName: '', productCode: '',
            mfgDate: '', expiry: '', qty: '',
            bufferSemiProduct: false,
            bufferUsageLedger: [],
            bufferLedgerInitialStock: 0,
            nonBufferLevel1: null,
            nonBufferPerformanceTestUsage: null,
            nonBufferInventoryAfterPerfTest: null,
            nonBufferLevel1LedgerRows: [],
            nonBufferLedgerInitialStock: null,
        });
        /** 비버퍼 반제품 관리 — 성능검사 행 사용일자(수기, 미저장) */
        const semiMgmtPerfTestDate = ref('');
        const semiMgmtContext = ref({ parentLot: '', semiLot: '', division: '' });

        const openHistorySemiManufacturingRecord = (semi) => {
            if (!historyDetail.value) {
                alert('먼저 Lot을 조회해 주세요.');
                return;
            }
            const semiLots = splitLines(semi.calcLot || semi.calc_lot || '');
            if (!semiLots.length) {
                alert('반제품 Lot 정보가 없습니다.');
                return;
            }
            const l2All = historyDetail.value.level2 || [];
            const l3All = historyDetail.value.level3 || [];
            const l2 = l2All.filter(r => {
                const pl = r['상위Lot'] || r['상위 Lot'] || '';
                const plParts = splitLines(pl);
                return semiLots.some(sl => pl === sl || plParts.includes(sl));
            });
            const onlyLevel3Materials = semiMrUsesLevel3Only(semi.division);
            const l2ChildLots = new Set();
            l2.forEach(r => splitLines(r['Lot No.'] || r['할당 Lot'] || '').forEach(x => l2ChildLots.add(x)));
            let l3 = [];
            if (onlyLevel3Materials) {
                /** Level3 상위Lot이 (1) L2 할당 Lot이거나, (2) 반제품 Lot을 직접 가리키는 경우 */
                const l3ParentMatchesSemiLot = (parentLotRaw) => {
                    const pl = String(parentLotRaw || '').trim();
                    if (!pl) return false;
                    const plParts = splitLines(pl);
                    return semiLots.some(sl => {
                        if (!sl) return false;
                        return pl === sl || plParts.includes(sl);
                    });
                };
                l3 = l3All.filter(r => {
                    const pl = r['상위Lot'] || r['상위 Lot'] || '';
                    const plTokens = splitLines(pl);
                    if (plTokens.some(t => t && l2ChildLots.has(t)) || l2ChildLots.has(pl)) return true;
                    if (l3ParentMatchesSemiLot(pl)) return true;
                    return false;
                });
            }
            const l0 = historyDetail.value.level0 || {};
            let productionQty = semi.productionQty ?? '';
            if (!onlyLevel3Materials) {
                const kit = l0['생산 수량(kit)'];
                const alt = l0.targetQty;
                if (kit !== undefined && kit !== null && String(kit).trim() !== '') productionQty = kit;
                else if (alt !== undefined && alt !== null && String(alt).trim() !== '') productionQty = alt;
            }
            semiMrPreview.value = {
                division: semi.division || '',
                instructionNo: semi.instructionNo || '',
                lotNo: semi.calcLot || semi.calc_lot || '',
                productionQty,
                mfgDate: semi.mfgDate || '',
                l2: onlyLevel3Materials ? [] : l2,
                l3,
                onlyLevel3Materials,
            };
            showSemiMrModal.value = true;
        };

        const openHistorySemiProductManagement = async (semi) => {
            const parent = selectedHistoryLot.value;
            if (!parent) {
                alert('Lot No.를 선택한 뒤 조회해 주세요.');
                return;
            }
            const semiLot = semi.calcLot || semi.calc_lot || '';
            const division = semi.division || '';
            try {
                const q = new URLSearchParams({ parent_lot: parent, semi_lot: semiLot, division });
                const res = await fetch(`/api/semi_product_management_preview?${q}`);
                const data = await res.json();
                if (data.error) {
                    alert(data.error);
                    return;
                }
                semiMgmtPreview.value = data;
                semiMgmtPerfTestDate.value = '';
                semiMgmtContext.value = { parentLot: parent, semiLot, division };
                showSemiMgmtModal.value = true;
            } catch (e) {
                alert('미리보기 로드 실패: ' + e);
            }
        };

        const openUploadSemiProductManagement = async (semi) => {
            const parent = csvLevel0.lotNo;
            if (!parent || !String(parent).trim()) {
                alert('Level 0의 LOT No.를 입력해 주세요.');
                return;
            }
            const semiLot = semi.calcLot || '';
            const division = semi.division || '';
            try {
                const q = new URLSearchParams({ parent_lot: parent.trim(), semi_lot: semiLot, division });
                const res = await fetch(`/api/semi_product_management_preview?${q}`);
                const data = await res.json();
                if (data.error) {
                    alert(data.error + '\n(저장된 BOM이 DB에 없으면 제조지시 기록에서 조회하거나 먼저 DB 저장을 하세요.)');
                    return;
                }
                semiMgmtPreview.value = data;
                semiMgmtPerfTestDate.value = '';
                semiMgmtContext.value = { parentLot: parent.trim(), semiLot, division };
                showSemiMgmtModal.value = true;
            } catch (e) {
                alert('미리보기 로드 실패: ' + e);
            }
        };

        const downloadSemiProductManagementFile = async (includeUsageHistory) => {
            const { parentLot, semiLot, division } = semiMgmtContext.value;
            if (!parentLot) return;
            const q = new URLSearchParams({
                parent_lot: parentLot,
                semi_lot: semiLot || '',
                division: division || '',
                include_usage_history: includeUsageHistory ? '1' : '0',
            });
            if (includeUsageHistory) {
                q.set('perf_test_date', semiMgmtPerfTestDate.value || '');
            }
            const url = `/api/semi_product_management_download?${q}`;
            try {
                const res = await fetch(url);
                if (!res.ok) {
                    let msg = '다운로드에 실패했습니다.';
                    try {
                        const j = await res.json();
                        if (j.error) msg = j.error;
                    } catch (e) { /* ignore */ }
                    alert(msg);
                    return;
                }
                const ct = (res.headers.get('content-type') || '').toLowerCase();
                if (ct.includes('json')) {
                    try {
                        const j = await res.json();
                        alert(j.error || '서버가 오류 응답을 반환했습니다.');
                    } catch (e) {
                        alert('서버가 엑셀이 아닌 응답을 반환했습니다.');
                    }
                    return;
                }
                const blob = await res.blob();
                const dispo = res.headers.get('content-disposition') || '';
                let fname = `Semi_Product_Management_${String(semiLot || semiMgmtPreview.value.N7 || parentLot).replace(/[^\w\-]+/g, '_').slice(0, 80)}${includeUsageHistory ? '_usage_history' : '_no_usage_history'}.xlsx`;
                const m = dispo.match(/filename\*?=(?:UTF-8'')?([^;\n]+)/i);
                if (m && m[1]) {
                    try {
                        fname = decodeURIComponent(m[1].trim().replace(/^["']|["']$/g, ''));
                    } catch (e) { /* keep default */ }
                }
                const objUrl = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = objUrl;
                link.download = fname;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(objUrl);
            } catch (e) {
                alert('다운로드 오류: ' + e);
            }
        };

        // 기록 조회용 하위 항목 필터링 (Level 2)
        const getHistoryL2 = (l1Lot) => {
            if (!historyDetail.value) return [];
            const lots = splitLines(l1Lot);
            const l2 = (historyDetail.value.level2 || []).filter(i => lots.includes(i['상위Lot'] || i['상위 Lot']));
            return groupRowsByLotAndCode(l2);
        };

        // 기록 조회용 하위 항목 필터링 (Level 3)
        const getHistoryL3 = (l2Lot) => {
            if (!historyDetail.value) return [];
            const lots = splitLines(l2Lot);
            const l3 = (historyDetail.value.level3 || []).filter(i => lots.includes(i['상위Lot'] || i['상위 Lot']));
            return groupRowsByLotAndCode(l3);
        };

        // =============================================
        // Lot 계층 구조 적용 로직
        // =============================================
        const applyLotHierarchy = () => {
            const rows = csvRows.value;
            if (!rows || rows.length === 0) return;

            // 유효기간 계산 (제조일자 + 1년 - 1일)
            if (csvLevel0.mfgDate) {
                try {
                    const mfg = new Date(csvLevel0.mfgDate);
                    if (!isNaN(mfg)) {
                        const exp = new Date(mfg);
                        const day = exp.getDate();
                        exp.setFullYear(exp.getFullYear() + 1);
                        if (exp.getDate() !== day) exp.setDate(0);
                        exp.setDate(exp.getDate() - 1);
                        csvLevel0.expiryDate = exp.toISOString().slice(0, 10);
                    }
                } catch(e) { console.error('Expiry calculation error:', e); }
            }

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
            
            // 완제품 패널(PI 항목) 업데이트
            fetchDocMaster(csvLevel0.modelName);
            
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
            const parts = filename.split('_');
            
            // 파일명 기반 자동 파싱 (모델명_BOM_제조수량_...)
            if (parts.length >= 1) {
                csvLevel0.modelName = parts[0];
                fetchItemMasterDetail(); 
            }
            if (parts.length >= 3) {
                csvLevel0.targetQty = parts[2];
            }
            
            calculateExpiry(); // 초기 유효기간 계산

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
        const resetCsv = () => { 
            csvRows.value = []; 
            csvFileName.value = ''; 
            // Level 0 입력값 초기화
            csvLevel0.productName = '';
            csvLevel0.modelName = '';
            csvLevel0.productInfo = '';
            csvLevel0.targetQty = '';
            csvLevel0.version = '';
            csvLevel0.mfgDate = new Date().toISOString().slice(0, 10);
            csvLevel0.requestTeam = '';
            csvLevel0.purpose = '';
            // 우측 패널 데이터 초기화
            docMasterList.value = [];
            semiLotList.value = [];
        };

        const downloadCsvResult = () => {
            if (!csvFiltered.value.length || typeof XLSX === 'undefined') return;
            const headers = ['Level', '상위 Lot', 'Code No.', '명칭 / 구성품', '필요 수량', '단위', '할당 Lot', '유효기간', '할당수량'];
            
            // Level 0 row construction (Production Info)
            const level0Row = [
                '0', 
                '', 
                csvLevel0.modelName, 
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
        const splitLines = (str) => {
            if (!str) return [];
            return String(str).split(/[\n,;]+/).map(s => s.trim()).filter(Boolean);
        };

        const groupItems = (items, keyCols, valCols, numCols = []) => {
            if (!items || !items.length) return [];
            const groups = new Map();
            items.forEach(item => {
                const key = keyCols.map(c => String(item[c] || '').trim()).join('||');
                if (!groups.has(key)) {
                    const g = { ...item };
                    valCols.forEach(c => { g[c] = String(item[c] || '').trim(); });
                    numCols.forEach(c => { 
                        const v = parseFloat(String(item[c] || '0').replace(/,/g, ''));
                        g[`_sum_${c}`] = isNaN(v) ? 0 : v;
                    });
                    groups.set(key, g);
                } else {
                    const g = groups.get(key);
                    valCols.forEach(c => { 
                        const v = String(item[c] || '').trim();
                        const currentV = String(g[c] || '').trim();
                        // 동일한 값이 이미 들어있지 않은 경우에만 추가 (Lot No. 등이 중복되지 않게)
                        if (v && !currentV.split('\n').includes(v)) {
                            g[c] += '\n' + v; 
                        }
                    });
                    numCols.forEach(c => {
                        const v = parseFloat(String(item[c] || '0').replace(/,/g, ''));
                        g[`_sum_${c}`] += isNaN(v) ? 0 : v;
                    });
                }
            });
            const result = Array.from(groups.values());
            result.forEach(g => {
                numCols.forEach(c => {
                    g[`_sum_${c}`] = Math.round(g[`_sum_${c}`] * 1000) / 1000;
                });
            });
            return result;
        };

        const filteredL1 = computed(() => {
            const l1 = viewData.value.level1.filter(i => i['상위Lot'] === selectedViewLot.value);
            return groupItems(l1, ['코드번호', '구성품 명칭'], ['Lot No.', '제조일자', '유효기간', '포장시 요구량'], ['포장시 요구량']);
        });

        const getL2SubItems = (l1LotStr) => {
            const lots = splitLines(l1LotStr);
            const l2 = viewData.value.level2.filter(i => lots.includes(i['상위Lot']));
            return groupItems(l2, ['상위Lot', '코드번호', '원재료명', '제조사'], ['Lot No.', '제조일자', '유효기간', '제조량'], ['제조량']);
        };

        const getL3SubItems = (l2Items) => {
            const l2Lots = l2Items.flatMap(i => splitLines(i['Lot No.']));
            const l3 = viewData.value.level3.filter(i => l2Lots.includes(i['상위Lot']));
            return groupItems(l3, ['상위Lot', '코드번호', '원재료명', '제조사'], ['Lot No.', '제조일자', '유효기간', '제조량'], ['제조량']);
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
            const day = d.getDate();
            d.setMonth(d.getMonth() + 13);
            if (d.getDate() !== day) d.setDate(0);
            d.setDate(d.getDate() - 1);
            const yyyy = d.getFullYear();
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            return `${yyyy}-${mm}-${dd}`;
        };

        const openSemiLotModal = async () => {
            if (!csvLevel0.modelName) {
                alert("먼저 Level 0의 모델명을 확인해주세요.");
                return;
            }
            try {
                const res = await fetch(`/api/doc_master/${encodeURIComponent(csvLevel0.modelName)}`);
                const data = await res.json();
                if (Array.isArray(data) && data.length) {
                    const filteredData = data.filter(d => {
                        const div = String(d.division || '').toUpperCase();
                        return !div.startsWith('LA') && !div.startsWith('PI');
                    });
                    semiLotList.value = filteredData.map(d => ({
                        ...d, mfgDate: '', calcLot: '', expiryDate: ''
                    }));
                    showSemiLotModal.value = true;
                } else {
                    alert('해당 모델명에 대한 반제품 제조지침서를 찾을 수 없습니다.');
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
            loadHistoryLots();
            calculateExpiry(); // 초기 오늘 날짜 기준 유효기간 설정
        });

        return {
            currentTab, breadcrumbLabel, loadViewerData,
            viewData, isViewLoading, selectedViewLot, currentL0, displayL0Info, selectViewLot,
            filteredL1, isOpen, toggleOpen, hasChildren, getL2SubItems, getL3SubItems, splitLines,
            showSemiModal, showRawModal, filteredInstructions, aggregatedMaterials,
            viewDepth, setViewDepth,
            // CSV Upload Tab
            miLatestLoading, miLatestError, miLatestResult, miManageFolder, selectedMiManageEntry, fetchManufacturingInstructionLatest,
            csvRows, csvFileName, isDragOver, csvInput, csvLevel0, csvFiltered, csvByLevel, csvByLevelGrouped,
            triggerCsvInput, handleCsvDrop, handleCsvFile, resetCsv, downloadCsvResult, isExpiryNear,
            showSemiLotModal, semiLotList, openSemiLotModal, onSemiMfgDateChange, applySemiLots,
            onModelNameChange, calculateExpiry, fetchItemMasterDetail,
            isCsvLevel0Valid, applyLotHierarchy,
            piList,
            saveToDatabase,
            // History Tab
            historyLots, selectedHistoryLot, historyDetail, loadHistoryLots, loadHistoryDetail, historyByLevelGrouped, historyPiList, historySemiLotList,
            getHistoryL2, getHistoryL3, historyDepth, setHistoryDepth,
            // New Buttons
            openPackagingInstruction, openProductManagement,
            showPackagingModal, packagingPreview, downloadPackagingFile,
            // 완제품 관리 관련
            showProductManagementModal, productManagementPreview, downloadProductManagementFile,
            // 제조지시 기록 — 반제품
            showSemiMrModal, semiMrPreview, openHistorySemiManufacturingRecord,
            showSemiMgmtModal, semiMgmtPreview, semiMgmtContext, semiMgmtPerfTestDate,
            openHistorySemiProductManagement, openUploadSemiProductManagement, downloadSemiProductManagementFile
        };
    }
}).mount('#app');
