// =====================================
// 네니아 완전 자동화 스크립트 Ultimate (투명도 제어 개선)
// 기본정보 + 인증마크 + 추가텍스트 + 보관방법 색상 + 개선된 투명도 제어
// =====================================

var DEBUG_MODE = false;
var TEMPLATE_LAYER_NAME = "auto_layer";

// 보관방법별 CMYK 색상
var STORAGE_COLORS = {
    "냉동": {c: 100, m: 40, y: 0, k: 0},
    "냉장": {c: 100, m: 0, y: 100, k: 0},
    "상온": {c: 0, m: 0, y: 20, k: 80}
};

// 인증마크 동적 배치 설정
var CERTIFICATION_LAYOUT = {
    START_OFFSET_X: 0,
    START_OFFSET_Y: 0,
    SPACING: 5,
    ROW_HEIGHT: 25
};

// 인증마크 목록
var CERTIFICATION_MARKS = [
    "HACCP", "무농약", "무농약가공", "유기농", "유기가공", "전통식품"
];

// 추가 텍스트 필드 목록
var ADDITIONAL_TEXT_FIELDS = [
    "자연해동", "D-7발주", "개별포장", "NEW"
];

function log(msg) {
    if (DEBUG_MODE) {
        $.writeln("[DEBUG] " + msg);
    }
}

// ✅ 안전한 trim 함수
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

// ✅ 통일된 그룹 찾기 함수 (중첩 그룹 지원)
function findGroupByNameUnified(name, container) {
    var patterns = [
        "<Group-" + name + ">",  // 모든 그룹이 이 패턴
        "Group-" + name,         
        name                     // 백업
    ];
    
    // 직접 그룹 검색
    for (var i = 0; i < container.groupItems.length; i++) {
        var group = container.groupItems[i];
        for (var p = 0; p < patterns.length; p++) {
            if (group.name === patterns[p]) {
                return group;
            }
        }
    }
    
    // 중첩 그룹 검색
    for (var j = 0; j < container.groupItems.length; j++) {
        var parentGroup = container.groupItems[j];
        var found = findGroupByNameUnified(name, parentGroup);
        if (found) {
            return found;
        }
    }
    
    return null;
}

// ✅ 텍스트 프레임 찾기
function findTextInLayer(name, layer) {
    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        if (item.typename === "TextFrame" && item.name === name) {
            return item;
        }
    }

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

// ✅ 연결된 이미지 찾기
function findPlacedItemByName(name, layer) {
    for (var i = 0; i < layer.placedItems.length; i++) {
        var placedItem = layer.placedItems[i];
        if (placedItem.name === name) {
            return placedItem;
        }
    }

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

    return searchPlacedItemsInGroups(layer);
}

// ✅ 혼합 검색 함수
function findItemByName(name, layer) {
    var textFrame = findTextInLayer(name, layer);
    if (textFrame) {
        return {type: 'TextFrame', item: textFrame};
    }

    var placedItem = findPlacedItemByName(name, layer);
    if (placedItem) {
        return {type: 'PlacedItem', item: placedItem};
    }

    return null;
}

// ✅ 개선된 추가 텍스트 그룹 제어 (투명도 방식)
function repositionAdditionalGroupsImproved(productIndex, productData, layer) {
    log('추가 텍스트 그룹 개선된 제어 시작: 상품 ' + productIndex);
    
    var activeGroups = [];
    
    // 1단계: 모든 추가 텍스트 그룹을 opacity 0으로 숨김
    for (var i = 0; i < ADDITIONAL_TEXT_FIELDS.length; i++) {
        var groupKey = ADDITIONAL_TEXT_FIELDS[i];
        var groupName = groupKey + productIndex;
        var group = findGroupByNameUnified(groupName, layer);
        
        if (group) {
            try {
                group.opacity = 0; // 모든 그룹 숨김
                log('그룹 숨김: ' + groupName + ' → opacity 0');
            } catch (e) {
                log('그룹 숨김 실패: ' + groupName + ' - ' + e.message);
            }
        }
    }
    
    // 2단계: 데이터가 있고 "Y" 값인 그룹만 opacity 100으로 표시
    for (var j = 0; j < ADDITIONAL_TEXT_FIELDS.length; j++) {
        var groupKey = ADDITIONAL_TEXT_FIELDS[j];
        var groupName = groupKey + productIndex;
        var group = findGroupByNameUnified(groupName, layer);
        var value = safeTrim(productData[groupKey] || '');
        
        if (group && value.toUpperCase() === 'Y') {
            try {
                group.opacity = 100; // Y값인 그룹만 표시
                
                if (groupKey !== 'NEW') { // NEW는 세로 정렬 안함
                    activeGroups.push({
                        name: groupName,
                        group: group,
                        data: value
                    });
                }
                
                log('✅ 그룹 표시: ' + groupName + ' → opacity 100 (Y값)');
            } catch (e) {
                log('그룹 표시 실패: ' + groupName + ' - ' + e.message);
            }
        }
    }
    
    // 3단계: 활성 그룹들을 세로로 정렬
    if (activeGroups.length > 0) {
        var baseGroup = activeGroups[0].group;
        var startX = baseGroup.position[0];
        var startY = baseGroup.position[1];
        var currentY = startY;
        var verticalSpacing = 5;
        
        for (var k = 0; k < activeGroups.length; k++) {
            var activeGroup = activeGroups[k];
            try {
                activeGroup.group.position = [startX, currentY];
                var groupHeight = getGroupHeight(activeGroup.group);
                currentY -= (groupHeight + verticalSpacing);
                log('✅ 그룹 세로 정렬: ' + activeGroup.name);
            } catch (e) {
                log('그룹 정렬 실패: ' + activeGroup.name + ' - ' + e.message);
            }
        }
    }
    
    log('추가 텍스트 그룹 제어 완료: ' + activeGroups.length + '개 활성화');
}

// ✅ 그룹 높이 계산
function getGroupHeight(group) {
    try {
        var bounds = group.geometricBounds;
        return Math.abs(bounds[1] - bounds[3]);
    } catch (e) {
        log('그룹 높이 계산 오류: ' + e.message);
        return 25;
    }
}

// ✅ 인증마크 동적 위치 조정
function repositionCertificationMarks(productIndex, productData, layer) {
    log('인증마크 동적 위치 조정: 상품 ' + productIndex);
    
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

    // 모든 인증마크 먼저 숨김
    for (var j = 0; j < allCertificationItems.length; j++) {
        allCertificationItems[j].result.item.opacity = 0;
    }

    if (activeCertifications.length === 0) {
        log('상품 ' + productIndex + ': 활성 인증마크 없음');
        return;
    }

    // 기준점 찾기
    var basePosition = getBaseCertificationPosition(productIndex, layer);
    if (!basePosition) {
        for (var m = 0; m < activeCertifications.length; m++) {
            activeCertifications[m].result.item.opacity = 100;
        }
        return;
    }

    // 활성 인증마크들을 연속으로 배치
    var currentX = basePosition.x + CERTIFICATION_LAYOUT.START_OFFSET_X;
    var currentY = basePosition.y + CERTIFICATION_LAYOUT.START_OFFSET_Y;
    
    for (var k = 0; k < activeCertifications.length; k++) {
        var cert = activeCertifications[k];
        var item = cert.result.item;
        
        try {
            item.position = [currentX, currentY];
        } catch (e) {
            log('위치 설정 오류: ' + cert.fullName + ' - ' + e.message);
        }

        item.opacity = 100;
        log('✅ 인증마크 동적 배치: ' + cert.fullName + ' → x=' + currentX + ', y=' + currentY);

        var itemWidth = getItemWidth(item, cert.result.type);
        currentX += itemWidth + CERTIFICATION_LAYOUT.SPACING;
    }

    log('상품 ' + productIndex + ' 인증마크 동적 배치 완료: ' + activeCertifications.length + '개');
}

function getBaseCertificationPosition(productIndex, layer) {
    var baseName = CERTIFICATION_MARKS[0] + productIndex;
    var baseResult = findItemByName(baseName, layer);
    if (baseResult) {
        try {
            return {
                x: baseResult.item.position[0],
                y: baseResult.item.position[1]
            };
        } catch (e) {
            log('기준점 위치 오류: ' + e.message);
        }
    }
    return null;
}

function getItemWidth(item, itemType) {
    try {
        var bounds;
        if (itemType === 'TextFrame') {
            bounds = item.visibleBounds;
        } else if (itemType === 'PlacedItem') {
            bounds = item.geometricBounds;
        } else {
            return 30;
        }
        return Math.abs(bounds[2] - bounds[0]);
    } catch (e) {
        log('너비 계산 오류: ' + e.message);
        return 30;
    }
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

// ✅ 완전 자동화 상품 업데이트 (모든 기능 통합)
function updateProductComplete(productIndex, productData, layer) {
    log("완전 자동화 업데이트 시작: 상품 " + productIndex);

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
            
            if (field.type === "storage" && field.value) {
                applyStorageColor(textFrame, field.value);
            }
            
            log('✅ 기본 필드: ' + field.name + ' = "' +
                (field.value.length > 30 ? field.value.substring(0, 30) + '...' : field.value) + '"');
        }
    }

    // 2. 인증마크 동적 위치 조정
    repositionCertificationMarks(productIndex, productData, layer);

    // 3. ✅ 개선된 추가 텍스트 그룹 제어 (투명도 방식)
    repositionAdditionalGroupsImproved(productIndex, productData, layer);

    log("완전 자동화 업데이트 완료: 상품 " + productIndex);
}

// CSV 파싱 함수들
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

function getAdditionalTextList(productData) {
    var additionalList = [];
    for (var i = 0; i < ADDITIONAL_TEXT_FIELDS.length; i++) {
        var field = ADDITIONAL_TEXT_FIELDS[i];
        var value = productData[field];
        if (value && safeTrim(value).length > 0) {
            additionalList.push(field + ": " + value);
        }
    }
    return additionalList;
}

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
    
    var pageOptions = "네니아 완전 자동화 - 사용 가능한 페이지:\n\n";
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

function getLayerByName(layerName, doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return doc.layers[i];
        }
    }
    return null;
}

// ✅ 메인 완전 자동화 함수 (개선된 투명도 제어 적용)
function applyNeniaCompleteAutomation() {
    try {
        log("네니아 완전 자동화 시작 (투명도 제어 개선)");
        
        var doc = app.activeDocument;
        var templateLayer = getLayerByName(TEMPLATE_LAYER_NAME, doc);
        
        if (!templateLayer) {
            alert('❌ 템플릿 레이어 "' + TEMPLATE_LAYER_NAME + '"를 찾을 수 없습니다.\n\n' +
                  '완전 자동화 설정 방법:\n' +
                  '1. 레이어 생성 및 이름을 "' + TEMPLATE_LAYER_NAME + '"로 설정\n' +
                  '2. 기본 텍스트 프레임: 페이지, 순서1~4, 상품명1~4, 용량1~4, 원재료1~4, 보관방법1~4\n' +
                  '3. 인증마크: HACCP1~4, 무농약1~4, 무농약가공1~4, 유기농1~4, 유기가공1~4, 전통식품1~4\n' +
                  '4. 추가 텍스트 그룹: <Group-자연해동1~4>, <Group-D-7발주1~4>, <Group-개별포장1~4>, <Group-NEW1~4>\n' +
                  '5. 스크립트 다시 실행');
            return;
        }

        var csvFile = File.openDialog("네니아 완전 자동화 CSV 파일 선택", "*.csv");
        if (!csvFile) {
            return;
        }

        csvFile.open("r");
        var content = csvFile.read();
        csvFile.close();
        
        if (!content || content.length === 0) {
            alert("파일이 비어있습니다.");
            return;
        }

        var lines = safeSplitLinesFixed(content);
        
        // ✅ 완전 확장 데이터 추출 (18개 컬럼)
        var allProductData = [];
        
        for (var i = 2; i < lines.length; i++) {
            var line = lines[i];
            
            if (typeof line === "string" && safeLength(safeTrim(line)) > 0) {
                var cells = parseCSVSafe(line);

                // ✅ 17개 이상 컬럼 확인 (수정된 조건)
                if (cells.length >= 17) {
                    allProductData.push({
                        // 기본 정보
                        페이지: parseInt(safeTrim(cells[0])) || 0,
                        순서: parseInt(safeTrim(cells[1])) || 0,
                        타이틀: cells[2] || '',
                        상품명: cells[3] || '',
                        용량: cells[4] || '',
                        원재료: cells[5] || '',
                        보관방법: cells[6] || '',
                        
                        // 인증마크 정보
                        HACCP: safeTrim(cells[7]).toUpperCase() === 'Y',
                        무농약: safeTrim(cells[8]).toUpperCase() === 'Y',
                        무농약가공: safeTrim(cells[9]).toUpperCase() === 'Y',
                        유기농: safeTrim(cells[10]).toUpperCase() === 'Y',
                        유기가공: safeTrim(cells[11]).toUpperCase() === 'Y',
                        전통식품: safeTrim(cells[12]).toUpperCase() === 'Y',
                        
                        // ✅ 추가 텍스트 정보 (투명도 제어용)
                        자연해동: cells[14] || '',
                        "D-7발주": cells[15] || '',
                        개별포장: cells[16] || '',
                        NEW: cells[17] || ''
                    });
                }
            }
        }

        if (allProductData.length === 0) {
            alert("상품 데이터가 없습니다!\n\n완전 자동화 CSV 형식 (17개 이상 컬럼):\n" +
                  "페이지,순서,타이틀,상품명,용량,원재료,보관방법,HACCP,무농약,무농약가공,유기농,유기가공,전통식품,무항생제,자연해동,D-7발주,개별포장,NEW");
            return;
        }

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

        // ✅ 완전 자동화 미리보기
        var completePreview = "네니아 완전 자동화 - 페이지 " + targetPage + ":\n\n";
        
        for (var j = 0; j < productDataArray.length; j++) {
            var product = productDataArray[j];
            var certList = getCertificationList(product);
            var additionalList = getAdditionalTextList(product);
            
            completePreview += "순서" + product.순서 + ": " + product.상품명 + "\n";
            completePreview += " 용량: " + product.용량 + "\n";
            completePreview += " 보관: " + product.보관방법 + " (색상 자동 적용)\n";
            completePreview += " 인증: " + (certList.length > 0 ? certList.join(", ") : "없음") + "\n";
            completePreview += " 추가: " + (additionalList.length > 0 ? additionalList.join(", ") : "없음") + "\n";
            completePreview += " 원재료: " + (product.원재료.length > 30 ? product.원재료.substring(0, 30) + "..." : product.원재료) + "\n\n";
        }

        if (confirm(completePreview + "위 내용으로 완전 자동화를 실행하시겠습니까?\n\n" +
                   "✅ 기본 정보 자동 적용\n" +
                   "✅ 보관방법별 색상 자동 적용\n" +
                   "✅ 인증마크 동적 위치 조정\n" +
                   "✅ 개선된 투명도 제어 (opacity 0→Y만 100)\n" +
                   "✅ 추가 텍스트 자동 적용 및 세로 정렬\n" +
                   "✅ CSV 콤마 포함 필드 완벽 처리")) {

            // ✅ 완전 자동화 실행
            for (var k = 0; k < productDataArray.length; k++) {
                var productNumber = k + 1;
                updateProductComplete(productNumber, productDataArray[k], templateLayer);
            }

            alert("🎉 네니아 완전 자동화 완료!\n\n" +
                  "페이지 " + targetPage + "의 " + productDataArray.length + "개 상품이 성공적으로 처리되었습니다.\n\n" +
                  "✅ 페이지/순서 처리 완료\n" +
                  "✅ 기본 정보 자동 적용 완료\n" +
                  "✅ 보관방법별 CMYK 색상 완료\n" +
                  "✅ 인증마크 동적 위치 조정 완료\n" +
                  "✅ 개선된 투명도 제어 완료 (opacity 0→Y만 100)\n" +
                  "✅ 추가 텍스트 자동 적용 및 세로 정렬 완료\n" +
                  "✅ 연결된 이미지 투명도 제어 완료\n" +
                  "✅ CSV 콤마 포함 원재료 완벽 처리 완료");
        }

    } catch (e) {
        alert("❌ 완전 자동화 오류: " + e.message + "\n라인: " + e.line);
        log("상세 오류: " + e.toString());
    }
}

// ✅ 네니아 완전 자동화 실행 (투명도 제어 개선 적용)
applyNeniaCompleteAutomation();
