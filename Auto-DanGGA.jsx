// =====================================
// 네니아 페이지별 정렬 처리 버전
// 페이지 번호 기준 필터링 및 정렬 적용
// =====================================

var DEBUG_MODE = false;
var TEMPLATE_LAYER_NAME = "auto_layer";

// ✅ 보관방법별 CMYK 색상 설정
var STORAGE_COLORS = {
    "냉동": {c: 100, m: 40, y: 0, k: 0},
    "냉장": {c: 100, m: 0, y: 100, k: 0},
    "상온": {c: 0, m: 0, y: 20, k: 80}
};

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

// ✅ 메인 함수 - 페이지별 정렬 처리
function applyNeniaPageSorted() {
    try {
        debugAlert("1. 네니아 페이지별 정렬 버전 시작");

        var doc = app.activeDocument;
        var templateLayer = getLayerByName(TEMPLATE_LAYER_NAME, doc);
        
        if (!templateLayer) {
            alert('❌ 템플릿 레이어 "' + TEMPLATE_LAYER_NAME + '"를 찾을 수 없습니다.');
            return;
        }

        var csvFile = File.openDialog("CSV 파일을 선택하세요", "*.csv");
        if (!csvFile) {
            debugAlert("파일 선택 취소");
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
        debugAlert("2. 총 줄 수: " + lines.length);

        // ✅ 모든 상품 데이터 먼저 추출
        var allProductData = [];
        
        for (var i = 2; i < lines.length; i++) {
            var line = lines[i];
            
            if (typeof line === "string" && safeLength(safeTrim(line)) > 0) {
                var cells = parseCSVSafe(line);

                if (cells.length >= 7) {
                    allProductData.push({
                        페이지: parseInt(safeTrim(cells[0])) || 0,
                        순서: parseInt(safeTrim(cells[1])) || 0,
                        타이틀: cells[2] || '',
                        상품명: cells[3] || '',
                        용량: cells[4] || '',
                        원재료: cells[5] || '',
                        보관방법: cells[6] || ''
                    });
                }
            }
        }

        if (allProductData.length === 0) {
            alert("상품 데이터가 없습니다!");
            return;
        }

        debugAlert("3. 전체 추출된 상품 수: " + allProductData.length);

        // ✅ 사용자에게 페이지 선택하게 하기
        var targetPage = getTargetPage(allProductData);
        if (targetPage === null) {
            return;
        }

        // ✅ 해당 페이지의 데이터만 필터링
        var pageProducts = [];
        for (var p = 0; p < allProductData.length; p++) {
            if (allProductData[p].페이지 === targetPage) {
                pageProducts.push(allProductData[p]);
            }
        }

        // ✅ 순서별로 정렬
        pageProducts.sort(function(a, b) {
            return a.순서 - b.순서;
        });

        // ✅ 최대 4개까지만 적용
        var productDataArray = pageProducts.slice(0, 4);

        debugAlert("4. 페이지 " + targetPage + " 상품 수: " + productDataArray.length);

        // 페이지 정보 표시
        var pageInfo = "페이지 " + targetPage + " 적용할 상품:\n";
        for (var j = 0; j < productDataArray.length; j++) {
            pageInfo += "순서" + productDataArray[j].순서 + ": " + productDataArray[j].상품명 + " (" + productDataArray[j].보관방법 + ")\n";
        }
        
        if (confirm(pageInfo + "\n위 상품들을 적용하시겠습니까?")) {
            // ✅ 정렬된 데이터 적용
            for (var k = 0; k < productDataArray.length; k++) {
                var productNumber = k + 1;
                debugAlert("5-" + productNumber + ". 상품 " + productNumber + " 적용 시작");
                updateExpandedProductData(productNumber, productDataArray[k], templateLayer);
            }

            alert("🎉 페이지별 정렬 완료!\n" + 
                  "페이지 " + targetPage + "의 " + productDataArray.length + "개 상품이 순서대로 적용되었습니다.\n\n" +
                  "✅ 페이지별 정렬 처리 완료\n" +
                  "✅ 순서별 정렬 완료\n" +
                  "✅ 보관방법별 색상 적용 완료");
        }

    } catch (e) {
        alert("❌ 오류: " + e.message + "\n라인: " + e.line);
        log("상세 오류: " + e.toString());
    }
}

// ✅ 사용자에게 페이지 선택하게 하는 함수
function getTargetPage(allData) {
    // 사용 가능한 페이지 목록 생성
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
    
    // 페이지 정렬
    pages.sort(function(a, b) { return a - b; });
    
    var pageOptions = "사용 가능한 페이지:\n\n";
    for (var j = 0; j < pages.length; j++) {
        pageOptions += "페이지 " + pages[j] + ": " + pageCount[pages[j]] + "개 상품\n";
    }
    
    var selectedPage = prompt(pageOptions + "\n적용할 페이지 번호를 입력하세요:");
    
    if (selectedPage === null) {
        return null;
    }
    
    var pageNum = parseInt(selectedPage);
    
    // 유효한 페이지인지 확인
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

// ✅ 확장된 상품 데이터 업데이트
function updateExpandedProductData(productIndex, productData, layer) {
    var fields = [
        {name: "페이지", value: productData.페이지.toString(), korean: "페이지", type: "normal"},
        {name: "순서" + productIndex, value: productData.순서.toString(), korean: "순서", type: "normal"},
        {name: "상품명" + productIndex, value: productData.상품명, korean: "상품명", type: "normal"},
        {name: "용량" + productIndex, value: productData.용량, korean: "용량", type: "normal"},
        {name: "원재료" + productIndex, value: productData.원재료, korean: "원재료", type: "normal"},
        {name: "보관방법" + productIndex, value: productData.보관방법, korean: "보관방법", type: "storage"}
    ];

    for (var f = 0; f < fields.length; f++) {
        var field = fields[f];
        var textFrame = findTextInLayer(field.name, layer);
        
        if (textFrame) {
            var oldContent = textFrame.contents;
            textFrame.contents = field.value;
            
            // 보관방법인 경우 색상 적용
            if (field.type === "storage" && field.value) {
                applyStorageColor(textFrame, field.value);
            }
            
            log('✅ 성공: ' + field.name + ' (' + field.korean + ') = "' + 
                (field.value.length > 30 ? field.value.substring(0, 30) + '...' : field.value) + '"');
        } else {
            log('❌ 실패: ' + field.name + ' 텍스트 프레임 없음');
        }
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
            log('⚠️ 알 수 없는 보관방법: "' + storage + '" (색상 적용 안됨)');
            
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

// ✅ 기존 함수들 (동일)
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

function findTextInLayer(name, layer) {
    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        if (item.typename === "TextFrame" && item.name === name) {
            return item;
        }
    }
    
    function searchGroups(container) {
        for (var j = 0; j < container.groupItems.length; j++) {
            var group = container.groupItems[j];
            
            for (var k = 0; k < group.textFrames.length; k++) {
                if (group.textFrames[k].name === name) {
                    return group.textFrames[k];
                }
            }
            
            var found = searchGroups(group);
            if (found) return found;
        }
        return null;
    }
    
    return searchGroups(layer);
}

function getLayerByName(layerName, doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return doc.layers[i];
        }
    }
    return null;
}

// ✅ 페이지별 정렬 버전 실행
applyNeniaPageSorted();
