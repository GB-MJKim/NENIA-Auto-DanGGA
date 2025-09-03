// =====================================
// 네니아 완전 통합 자동화 스크립트 Ultimate Version
// 페이지/순서 + 보관방법 색상 + 인증마크 동적 위치 조정 + CSV 콤마 처리
// =====================================

var DEBUG_MODE = false;
var TEMPLATE_LAYER_NAME = "auto_layer";

// ✅ 인증마크 동적 배치 설정
var CERTIFICATION_LAYOUT = {
    START_OFFSET_X: 0,      // 시작점에서 X축 오프셋
    START_OFFSET_Y: 0,      // 시작점에서 Y축 오프셋  
    SPACING: 5,             // 인증마크 간 간격 (pt)
    ROW_HEIGHT: 25          // 줄 높이 (여러 줄일 경우)
};

// ✅ 보관방법별 CMYK 색상
var STORAGE_COLORS = {
    "냉동": {c: 100, m: 40, y: 0, k: 0},
    "냉장": {c: 100, m: 0, y: 100, k: 0},
    "상온": {c: 0, m: 0, y: 20, k: 80}
};

// ✅ 인증마크 목록 (순서 중요!)
var CERTIFICATION_MARKS = [
    "HACCP",
    "무농약", 
    "무농약가공",
    "유기농",
    "유기가공", 
    "전통식품"
];

function log(msg) {
    if (DEBUG_MODE) {
        $.writeln("[DEBUG] " + msg);
    }
}

function debugAlert(msg) {
    if (DEBUG_MODE) {
        alert("[DEBUG] " + msg);
    }
}

// ✅ 안전한 trim 함수 (line.trim() 에러 방지)
function safeTrim(value) {
    if (typeof value !== "string") {
        return "";
    }
    
    if (typeof value.trim === "function") {
        return value.trim();
    }
    
    var startIdx = 0;
    var endIdx = value.length - 1;
    
    while (startIdx < value.length && 
           (value.charAt(startIdx) === " " || value.charAt(startIdx) === "\t" || 
            value.charAt(startIdx) === "\n" || value.charAt(startIdx) === "\r")) {
        startIdx++;
    }
    
    while (endIdx >= startIdx && 
           (value.charAt(endIdx) === " " || value.charAt(endIdx) === "\t" || 
            value.charAt(endIdx) === "\n" || value.charAt(endIdx) === "\r")) {
        endIdx--;
    }
    
    if (startIdx > endIdx) {
        return "";
    }
    
    var result = "";
    for (var i = startIdx; i <= endIdx; i++) {
        result += value.charAt(i);
    }
    
    return result;
}

function safeLength(value) {
    if (typeof value !== "string") {
        return 0;
    }
    return value.length;
}

// ✅ 텍스트 프레임 찾기
function findTextInLayer(name, layer) {
    // 레이어 내 직접 TextFrame
    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        if (item.typename === "TextFrame" && item.name === name) {
            return item;
        }
    }
    
    // 그룹 내 TextFrame 검색
    function searchTextInGroups(container) {
        for (var j = 0; j < container.groupItems.length; j++) {
            var group = container.groupItems[j];
            
            for (var k = 0; k < group.textFrames.length; k++) {
                if (group.textFrames[k].name === name) {
                    return group.textFrames[k];
                }
            }
            
            var found = searchTextInGroups(group);
            if (found) return found;
        }
        return null;
    }
    
    return searchTextInGroups(layer);
}

// ✅ 연결된 이미지(PlacedItem) 찾기
function findPlacedItemByName(name, layer) {
    // 레이어 내 직접 PlacedItems
    for (var i = 0; i < layer.placedItems.length; i++) {
        var placedItem = layer.placedItems[i];
        if (placedItem.name === name) {
            log('PlacedItem 발견: ' + name);
            return placedItem;
        }
    }
    
    // 그룹 내 PlacedItems 검색
    function searchPlacedItemsInGroups(container) {
        for (var j = 0; j < container.groupItems.length; j++) {
            var group = container.groupItems[j];
            
            for (var k = 0; k < group.placedItems.length; k++) {
                if (group.placedItems[k].name === name) {
                    return group.placedItems[k];
                }
            }
            
            var found = searchPlacedItemsInGroups(group);
            if (found) return found;
        }
        return null;
    }
    
    var groupResult = searchPlacedItemsInGroups(layer);
    if (groupResult) {
        log('그룹 내 PlacedItem 발견: ' + name);
    }
    
    return groupResult;
}

// ✅ 혼합 검색 함수 (TextFrame + PlacedItem)
function findItemByName(name, layer) {
    // 먼저 TextFrame 검색
    var textFrame = findTextInLayer(name, layer);
    if (textFrame) {
        return {type: 'TextFrame', item: textFrame};
    }
    
    // PlacedItem 검색
    var placedItem = findPlacedItemByName(name, layer);
    if (placedItem) {
        return {type: 'PlacedItem', item: placedItem};
    }
    
    return null;
}

// ✅ 메인 함수 - 완전 통합 Ultimate Version
function applyNeniaUltimate() {
    try {
        debugAlert("1. 네니아 완전 통합 Ultimate 버전 시작");

        var doc = app.activeDocument;
        var templateLayer = getLayerByName(TEMPLATE_LAYER_NAME, doc);
        
        if (!templateLayer) {
            alert('❌ 템플릿 레이어 "' + TEMPLATE_LAYER_NAME + '"를 찾을 수 없습니다.\n\n' +
                  '완전 자동화 설정 방법:\n' +
                  '1. 레이어 생성 및 이름을 "' + TEMPLATE_LAYER_NAME + '"로 설정\n' +
                  '2. 기본 텍스트 프레임: 페이지, 순서1~4, 상품명1~4, 용량1~4, 원재료1~4, 보관방법1~4\n' +
                  '3. 인증마크 (TextFrame 또는 PlacedItem): HACCP1~4, 무농약1~4, 무농약가공1~4, 유기농1~4, 유기가공1~4, 전통식품1~4\n' +
                  '4. 모든 인증마크를 투명도 0%로 설정\n' +
                  '5. 스크립트 다시 실행');
            return;
        }

        var csvFile = File.openDialog("네니아 완전 통합 CSV 파일 선택", "*.csv");
        if (!csvFile) {
            debugAlert("파일 선택 취소");
            return;
        }

        debugAlert("2. 파일 선택: " + csvFile.name);

        csvFile.open("r");
        var content = csvFile.read();
        csvFile.close();

        if (!content || content.length === 0) {
            alert("파일이 비어있습니다.");
            return;
        }

        var lines = safeSplitLinesFixed(content);
        debugAlert("3. 총 줄 수: " + lines.length);

        // ✅ 완전 통합 데이터 추출
        var allProductData = [];
        
        for (var i = 2; i < lines.length; i++) {
            var line = lines[i];
            
            if (typeof line === "string" && safeLength(safeTrim(line)) > 0) {
                var cells = parseCSVSafe(line);

                debugAlert("4-" + i + ". 파싱 결과: " + cells.length + "개 셀");
                
                // ✅ 13개 컬럼 확인 (7개 기본 + 6개 인증)
                if (cells.length >= 13) {
                    allProductData.push({
                        페이지: parseInt(safeTrim(cells[0])) || 0,
                        순서: parseInt(safeTrim(cells[1])) || 0,
                        타이틀: cells[2] || '',
                        상품명: cells[3] || '',
                        용량: cells[4] || '',
                        원재료: cells[5] || '',
                        보관방법: cells[6] || '',
                        // ✅ 인증마크 정보
                        HACCP: safeTrim(cells[7]).toUpperCase() === 'Y',
                        무농약: safeTrim(cells[8]).toUpperCase() === 'Y',
                        무농약가공: safeTrim(cells[9]).toUpperCase() === 'Y',
                        유기농: safeTrim(cells[10]).toUpperCase() === 'Y',
                        유기가공: safeTrim(cells[11]).toUpperCase() === 'Y',
                        전통식품: safeTrim(cells[12]).toUpperCase() === 'Y'
                    });
                    
                    if (DEBUG_MODE) {
                        var product = allProductData[allProductData.length - 1];
                        debugAlert("추가된 상품:\n" +
                                  "페이지: " + product.페이지 + ", 순서: " + product.순서 + "\n" +
                                  "상품명: " + product.상품명 + "\n" +
                                  "보관방법: " + product.보관방법 + "\n" +
                                  "인증: HACCP=" + product.HACCP + ", 무농약=" + product.무농약 + ", 유기농=" + product.유기농);
                    }
                }
            }
        }

        if (allProductData.length === 0) {
            alert("상품 데이터가 없습니다!\n\n완전 통합 CSV 형식:\n" +
                  "페이지,순서,타이틀,상품명,용량,원재료,보관방법,HACCP,무농약,무농약가공,유기농,유기가공,전통식품");
            return;
        }

        debugAlert("5. 전체 추출된 상품 수: " + allProductData.length);

        // 페이지 선택
        var targetPage = getTargetPage(allProductData);
        if (targetPage === null) {
            return;
        }

        // 해당 페이지 데이터 필터링 및 정렬
        var pageProducts = [];
        for (var p = 0; p < allProductData.length; p++) {
            if (allProductData[p].페이지 === targetPage) {
                pageProducts.push(allProductData[p]);
            }
        }

        pageProducts.sort(function(a, b) {
            return a.순서 - b.순서;
        });

        var productDataArray = pageProducts.slice(0, 4);

        debugAlert("6. 페이지 " + targetPage + " 상품 수: " + productDataArray.length);

        // ✅ 완전 통합 미리보기
        var ultimatePreview = "네니아 완전 통합 - 페이지 " + targetPage + ":\n\n";
        for (var j = 0; j < productDataArray.length; j++) {
            var product = productDataArray[j];
            var certList = getCertificationList(product);
            ultimatePreview += "순서" + product.순서 + ": " + product.상품명 + "\n";
            ultimatePreview += "  용량: " + product.용량 + "\n";
            ultimatePreview += "  보관: " + product.보관방법 + " (색상 자동 적용)\n";
            ultimatePreview += "  인증: " + (certList.length > 0 ? certList.join(", ") : "없음") + "\n";
            ultimatePreview += "  원재료: " + (product.원재료.length > 40 ? product.원재료.substring(0, 40) + "..." : product.원재료) + "\n\n";
        }
        
        if (confirm(ultimatePreview + "위 내용으로 완전 통합 자동화를 실행하시겠습니까?\n\n" +
                   "✅ 기본 정보 자동 적용\n" +
                   "✅ 보관방법별 색상 자동 적용\n" +
                   "✅ 인증마크 동적 위치 조정\n" +
                   "✅ CSV 콤마 포함 필드 완벽 처리")) {
            
            // ✅ 완전 통합 자동화 실행
            for (var k = 0; k < productDataArray.length; k++) {
                var productNumber = k + 1;
                debugAlert("7-" + productNumber + ". 상품 " + productNumber + " 완전 통합 처리 시작");
                updateProductUltimate(productNumber, productDataArray[k], templateLayer);
            }

            alert("🎉 네니아 완전 통합 자동화 완료!\n\n" + 
                  "페이지 " + targetPage + "의 " + productDataArray.length + "개 상품이 성공적으로 처리되었습니다.\n\n" +
                  "✅ 페이지/순서 처리 완료\n" +
                  "✅ 기본 정보 자동 적용 완료\n" +
                  "✅ 보관방법별 CMYK 색상 완료\n" +
                  "✅ 인증마크 동적 위치 조정 완료\n" +
                  "✅ 연결된 이미지 투명도 제어 완료\n" +
                  "✅ CSV 콤마 포함 원재료 완벽 처리 완료\n" +
                  "✅ 80페이지 단가책자 대응 완료");
        }

    } catch (e) {
        alert("❌ 완전 통합 오류: " + e.message + "\n라인: " + e.line);
        log("상세 오류: " + e.toString());
    }
}

// ✅ 완전 통합 상품 업데이트 (모든 기능 포함)
function updateProductUltimate(productIndex, productData, layer) {
    debugAlert("완전 통합 업데이트 시작: 상품 " + productIndex);
    
    // 1. 기본 필드 업데이트
    var basicFields = [
        {name: "페이지", value: productData.페이지.toString(), korean: "페이지", type: "normal"},
        {name: "순서" + productIndex, value: productData.순서.toString(), korean: "순서", type: "normal"},
        {name: "상품명" + productIndex, value: productData.상품명, korean: "상품명", type: "normal"},
        {name: "용량" + productIndex, value: productData.용량, korean: "용량", type: "normal"},
        {name: "원재료" + productIndex, value: productData.원재료, korean: "원재료", type: "normal"},
        {name: "보관방법" + productIndex, value: productData.보관방법, korean: "보관방법", type: "storage"}
    ];

    for (var f = 0; f < basicFields.length; f++) {
        var field = basicFields[f];
        var textFrame = findTextInLayer(field.name, layer);
        
        if (textFrame) {
            textFrame.contents = field.value;
            
            // ✅ 보관방법 색상 적용
            if (field.type === "storage" && field.value) {
                applyStorageColor(textFrame, field.value);
            }
            
            log('✅ 기본 필드: ' + field.name + ' = "' + 
                (field.value.length > 30 ? field.value.substring(0, 30) + '...' : field.value) + '"');
        } else {
            log('❌ 기본 필드 없음: ' + field.name);
        }
    }

    // 2. 인증마크 동적 위치 조정 및 적용
    repositionCertificationMarks(productIndex, productData, layer);
    
    debugAlert("완전 통합 업데이트 완료: 상품 " + productIndex);
}

// ✅ 인증마크 동적 위치 조정
function repositionCertificationMarks(productIndex, productData, layer) {
    debugAlert('인증마크 동적 위치 조정: 상품 ' + productIndex);
    
    // 1. 활성화된 인증마크 목록 생성
    var activeCertifications = [];
    var allCertificationItems = [];
    
    for (var i = 0; i < CERTIFICATION_MARKS.length; i++) {
        var mark = CERTIFICATION_MARKS[i];
        var markName = mark + productIndex;
        var result = findItemByName(markName, layer);
        
        if (result) {
            allCertificationItems.push({
                name: mark,
                fullName: markName,
                result: result,
                active: productData[mark]
            });
            
            if (productData[mark]) {
                activeCertifications.push({
                    name: mark,
                    fullName: markName,
                    result: result
                });
            }
        }
    }
    
    // 2. 모든 인증마크 먼저 숨김
    for (var j = 0; j < allCertificationItems.length; j++) {
        allCertificationItems[j].result.item.opacity = 0;
        log('인증마크 초기화: ' + allCertificationItems[j].fullName);
    }
    
    if (activeCertifications.length === 0) {
        debugAlert('상품 ' + productIndex + ': 활성 인증마크 없음');
        return;
    }
    
    // 3. 기준점 찾기
    var basePosition = getBaseCertificationPosition(productIndex, layer);
    if (!basePosition) {
        log('❌ 기준점을 찾을 수 없음: 상품 ' + productIndex);
        // 기준점이 없으면 고정 위치에 표시
        for (var m = 0; m < activeCertifications.length; m++) {
            activeCertifications[m].result.item.opacity = 100;
        }
        return;
    }
    
    debugAlert('기준점 설정: x=' + basePosition.x + ', y=' + basePosition.y);
    
    // 4. 활성 인증마크들을 연속으로 배치
    var currentX = basePosition.x + CERTIFICATION_LAYOUT.START_OFFSET_X;
    var currentY = basePosition.y + CERTIFICATION_LAYOUT.START_OFFSET_Y;
    
    for (var k = 0; k < activeCertifications.length; k++) {
        var cert = activeCertifications[k];
        var item = cert.result.item;
        
        // 위치 설정
        try {
            if (cert.result.type === 'TextFrame' || cert.result.type === 'PlacedItem') {
                item.position = [currentX, currentY];
            }
        } catch (e) {
            log('위치 설정 오류: ' + cert.fullName + ' - ' + e.message);
        }
        
        // 투명도 100%로 표시
        item.opacity = 100;
        
        log('✅ 인증마크 동적 배치: ' + cert.fullName + ' → x=' + currentX + ', y=' + currentY);
        
        // 다음 위치 계산
        var itemWidth = getItemWidth(item, cert.result.type);
        currentX += itemWidth + CERTIFICATION_LAYOUT.SPACING;
    }
    
    debugAlert('상품 ' + productIndex + ' 인증마크 동적 배치 완료: ' + activeCertifications.length + '개');
}

// ✅ 기준점 위치 찾기
function getBaseCertificationPosition(productIndex, layer) {
    var baseName = CERTIFICATION_MARKS[0] + productIndex; // HACCP1, HACCP2 등
    var baseResult = findItemByName(baseName, layer);
    
    if (baseResult) {
        var baseItem = baseResult.item;
        
        try {
            return {
                x: baseItem.position[0],
                y: baseItem.position[1]
            };
        } catch (e) {
            log('기준점 위치 오류: ' + e.message);
        }
    }
    
    return null;
}

// ✅ 아이템 너비 구하기
function getItemWidth(item, itemType) {
    try {
        var bounds;
        if (itemType === 'TextFrame') {
            bounds = item.visibleBounds;
        } else if (itemType === 'PlacedItem') {
            bounds = item.geometricBounds;
        } else {
            return 30; // 기본값
        }
        
        return Math.abs(bounds[2] - bounds[0]); // right - left
    } catch (e) {
        log('너비 계산 오류: ' + e.message);
        return 30; // 기본값
    }
}

// ✅ 상품별 인증 목록 생성
function getCertificationList(productData) {
    var certList = [];
    
    for (var i = 0; i < CERTIFICATION_MARKS.length; i++) {
        var mark = CERTIFICATION_MARKS[i];
        if (productData[mark]) {
            certList.push(mark);
        }
    }
    
    return certList;
}

// ✅ 페이지 선택 함수
function getTargetPage(allData) {
    var pages = [];
    var pageCount = {};
    
    for (var i = 0; i < allData.length; i++) {
        var page = allData[i].페이지;
        if (pageCount[page]) {
            pageCount[page]++;
        } else {
            pageCount[page] = 1;
            pages.push(page);
        }
    }
    
    pages.sort(function(a, b) { return a - b; });
    
    var pageOptions = "네니아 완전 통합 - 사용 가능한 페이지:\n\n";
    for (var j = 0; j < pages.length; j++) {
        pageOptions += "페이지 " + pages[j] + ": " + pageCount[pages[j]] + "개 상품\n";
    }
    
    var selectedPage = prompt(pageOptions + "\n적용할 페이지 번호를 입력하세요:");
    
    if (selectedPage === null) {
        return null;
    }
    
    var pageNum = parseInt(selectedPage);
    
    var isValidPage = false;
    for (var k = 0; k < pages.length; k++) {
        if (pages[k] === pageNum) {
            isValidPage = true;
            break;
        }
    }
    
    if (!isValidPage) {
        alert("❌ 잘못된 페이지 번호입니다: " + selectedPage);
        return null;
    }
    
    return pageNum;
}

// ✅ 보관방법별 색상 적용
function applyStorageColor(textFrame, storageMethod) {
    try {
        var storage = safeTrim(storageMethod);
        var colorData = STORAGE_COLORS[storage];
        
        if (colorData) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = colorData.c;
            cmykColor.magenta = colorData.m;
            cmykColor.yellow = colorData.y;
            cmykColor.black = colorData.k;
            
            var textRange = textFrame.textRange;
            textRange.characterAttributes.fillColor = cmykColor;
            
            log('✅ 색상 적용: ' + storage + ' → CMYK(' + 
                colorData.c + ',' + colorData.m + ',' + colorData.y + ',' + colorData.k + ')');
        } else {
            log('⚠️ 알 수 없는 보관방법: "' + storage + '" (기본 색상 적용)');
            
            var blackColor = new CMYKColor();
            blackColor.cyan = 0;
            blackColor.magenta = 0;
            blackColor.yellow = 0;
            blackColor.black = 100;
            
            var textRange = textFrame.textRange;
            textRange.characterAttributes.fillColor = blackColor;
        }
    } catch (e) {
        log('❌ 색상 적용 오류: ' + e.message);
    }
}

// ✅ 안전한 줄 분할 (콤마 포함 필드 지원)
function safeSplitLinesFixed(content) {
    var result = [];
    var current = "";
    var inQuotes = false;
    
    for (var i = 0; i < content.length; i++) {
        var currentChar = content.charAt(i);
        
        if (currentChar === '"') {
            inQuotes = !inQuotes;
            current += currentChar;
        } else if ((currentChar === "\n" || currentChar === "\r") && !inQuotes) {
            if (safeLength(safeTrim(current)) > 0) {
                result.push(current);
            }
            current = "";
            
            if (currentChar === "\r" && i + 1 < content.length && content.charAt(i + 1) === "\n") {
                i++;
            }
        } else {
            current += currentChar;
        }
    }
    
    if (safeLength(safeTrim(current)) > 0) {
        result.push(current);
    }
    
    return result;
}

// ✅ CSV 파싱 (콤마 포함 필드 완벽 처리)
function parseCSVSafe(line) {
    var result = [];
    var current = "";
    var inQuotes = false;
    var i = 0;
    
    while (i < line.length) {
        var currentChar = line.charAt(i);
        
        if (currentChar === '"') {
            if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"') {
                current += '"';
                i += 2;
                continue;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (currentChar === "," && !inQuotes) {
            result.push(cleanCellSafe(current));
            current = "";
        } else {
            current += currentChar;
        }
        
        i++;
    }
    
    result.push(cleanCellSafe(current));
    return result;
}

function cleanCellSafe(value) {
    if (typeof value !== "string") {
        return "";
    }
    
    value = safeTrim(value);
    
    if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
        value = value.substring(1, value.length - 1);
        value = value.replace(/""/g, '"');
    }
    
    return value;
}

function getLayerByName(layerName, doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return doc.layers[i];
        }
    }
    return null;
}

// ✅ 디버깅용 완전 테스트 함수
function testUltimateSetup() {
    var doc = app.activeDocument;
    var layer = getLayerByName(TEMPLATE_LAYER_NAME, doc);
    
    if (!layer) {
        alert("❌ 레이어 '" + TEMPLATE_LAYER_NAME + "' 없음");
        return;
    }
    
    var testInfo = "네니아 완전 통합 테스트 결과:\n\n";
    
    // 기본 필드 테스트
    testInfo += "=== 기본 필드 ===\n";
    var basicFields = ["페이지", "순서1", "상품명1", "용량1", "원재료1", "보관방법1"];
    for (var i = 0; i < basicFields.length; i++) {
        var tf = findTextInLayer(basicFields[i], layer);
        testInfo += basicFields[i] + ": " + (tf ? "✅" : "❌") + "\n";
    }
    
    // 인증마크 테스트
    testInfo += "\n=== 인증마크 (상품1) ===\n";
    for (var j = 0; j < CERTIFICATION_MARKS.length; j++) {
        var certName = CERTIFICATION_MARKS[j] + "1";
        var result = findItemByName(certName, layer);
        testInfo += certName + ": " + (result ? "✅ (" + result.type + ")" : "❌") + "\n";
    }
    
    alert(testInfo);
}

// ✅ 네니아 완전 통합 Ultimate 버전 실행
applyNeniaUltimate();

// 디버깅이 필요하면: testUltimateSetup();
