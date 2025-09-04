// =====================================
// 텍스트 입력 기반 요소 위치 확인 스크립트
// Group, Text, PlacedItem 모든 타입 지원
// =====================================

// ===== CONFIG =====
var CONFIG = {
    SEARCH_PATTERNS: {
        EXACT_MATCH: true,      // 정확히 일치
        PARTIAL_MATCH: true,    // 부분 일치
        CASE_SENSITIVE: false   // 대소문자 구분
    }
};

// ===== UTILS =====
var Utils = {
    log: function(message) {
        $.writeln("[SEARCH] " + message);
    },
    
    safeTrim: function(str) {
        if (typeof str !== "string") return "";
        return str.replace(/^\s+|\s+$/g, '');
    },
    
    formatPosition: function(position) {
        if (!position || position.length < 2) return "위치 없음";
        return "X: " + Math.round(position[0] * 100) / 100 + ", Y: " + Math.round(position[1] * 100) / 100;
    },
    
    formatBounds: function(bounds) {
        if (!bounds || bounds.length < 4) return "경계 없음";
        return "상: " + Math.round(bounds[1] * 100) / 100 + 
               ", 좌: " + Math.round(bounds[0] * 100) / 100 + 
               ", 하: " + Math.round(bounds[3] * 100) / 100 + 
               ", 우: " + Math.round(bounds[2] * 100) / 100;
    },
    
    matchText: function(searchText, targetText, targetName) {
        if (!searchText) return false;
        
        var search = CONFIG.SEARCH_PATTERNS.CASE_SENSITIVE ? searchText : searchText.toLowerCase();
        var target = CONFIG.SEARCH_PATTERNS.CASE_SENSITIVE ? targetText : targetText.toLowerCase();
        var name = CONFIG.SEARCH_PATTERNS.CASE_SENSITIVE ? targetName : targetName.toLowerCase();
        
        // 이름 일치 확인
        if (CONFIG.SEARCH_PATTERNS.EXACT_MATCH && name === search) return true;
        if (CONFIG.SEARCH_PATTERNS.PARTIAL_MATCH && name.indexOf(search) !== -1) return true;
        
        // 내용 일치 확인 (TextFrame의 경우)
        if (target) {
            if (CONFIG.SEARCH_PATTERNS.EXACT_MATCH && target === search) return true;
            if (CONFIG.SEARCH_PATTERNS.PARTIAL_MATCH && target.indexOf(search) !== -1) return true;
        }
        
        return false;
    }
};

// ===== ELEMENT FINDER =====
var ElementFinder = {
    foundElements: [],
    
    // ✅ 모든 요소 타입 통합 검색
    searchAllElements: function(container, searchText, containerName) {
        if (!container || !searchText) return;
        
        containerName = containerName || "Unknown";
        Utils.log("컨테이너 검색: " + containerName);
        
        // 1. GroupItem 검색
        this.searchGroupItems(container, searchText, containerName);
        
        // 2. TextFrame 검색
        this.searchTextFrames(container, searchText, containerName);
        
        // 3. PlacedItem 검색
        this.searchPlacedItems(container, searchText, containerName);
        
        // 4. PathItem 검색
        this.searchPathItems(container, searchText, containerName);
        
        // 5. 중첩 컨테이너 검색
        if (container.groupItems) {
            for (var i = 0; i < container.groupItems.length; i++) {
                this.searchAllElements(container.groupItems[i], searchText, 
                    containerName + " > " + container.groupItems[i].name);
            }
        }
    },
    
    // ✅ GroupItem 검색
    searchGroupItems: function(container, searchText, containerName) {
        if (!container.groupItems) return;
        
        for (var i = 0; i < container.groupItems.length; i++) {
            var group = container.groupItems[i];
            var groupName = Utils.safeTrim(group.name) || ("Group_" + i);
            
            if (Utils.matchText(searchText, "", groupName)) {
                try {
                    this.foundElements.push({
                        name: groupName,
                        type: "GroupItem",
                        position: group.position,
                        bounds: group.geometricBounds,
                        container: containerName,
                        opacity: group.opacity || 100,
                        visible: !group.hidden,
                        locked: group.locked || false
                    });
                    
                    Utils.log("GroupItem 발견: " + groupName + " 위치: " + Utils.formatPosition(group.position));
                } catch (e) {
                    Utils.log("GroupItem 정보 추출 실패: " + groupName + " - " + e.message);
                }
            }
        }
    },
    
    // ✅ TextFrame 검색
    searchTextFrames: function(container, searchText, containerName) {
        if (!container.textFrames) return;
        
        for (var i = 0; i < container.textFrames.length; i++) {
            var textFrame = container.textFrames[i];
            var frameName = Utils.safeTrim(textFrame.name) || ("Text_" + i);
            var frameContent = Utils.safeTrim(textFrame.contents) || "";
            
            if (Utils.matchText(searchText, frameContent, frameName)) {
                try {
                    this.foundElements.push({
                        name: frameName,
                        type: "TextFrame",
                        content: frameContent,
                        position: textFrame.position,
                        bounds: textFrame.geometricBounds,
                        container: containerName,
                        opacity: textFrame.opacity || 100,
                        visible: !textFrame.hidden,
                        locked: textFrame.locked || false
                    });
                    
                    Utils.log("TextFrame 발견: " + frameName + " 내용: '" + frameContent + "' 위치: " + Utils.formatPosition(textFrame.position));
                } catch (e) {
                    Utils.log("TextFrame 정보 추출 실패: " + frameName + " - " + e.message);
                }
            }
        }
    },
    
    // ✅ PlacedItem 검색
    searchPlacedItems: function(container, searchText, containerName) {
        if (!container.placedItems) return;
        
        for (var i = 0; i < container.placedItems.length; i++) {
            var placedItem = container.placedItems[i];
            var itemName = Utils.safeTrim(placedItem.name) || ("PlacedItem_" + i);
            
            if (Utils.matchText(searchText, "", itemName)) {
                try {
                    var fileInfo = "";
                    try {
                        fileInfo = placedItem.file ? placedItem.file.name : "파일 정보 없음";
                    } catch (e) {
                        fileInfo = "파일 접근 불가";
                    }
                    
                    this.foundElements.push({
                        name: itemName,
                        type: "PlacedItem",
                        fileName: fileInfo,
                        position: placedItem.position,
                        bounds: placedItem.geometricBounds,
                        container: containerName,
                        opacity: placedItem.opacity || 100,
                        visible: !placedItem.hidden,
                        locked: placedItem.locked || false
                    });
                    
                    Utils.log("PlacedItem 발견: " + itemName + " 파일: " + fileInfo + " 위치: " + Utils.formatPosition(placedItem.position));
                } catch (e) {
                    Utils.log("PlacedItem 정보 추출 실패: " + itemName + " - " + e.message);
                }
            }
        }
    },
    
    // ✅ PathItem 검색
    searchPathItems: function(container, searchText, containerName) {
        if (!container.pathItems) return;
        
        for (var i = 0; i < container.pathItems.length; i++) {
            var pathItem = container.pathItems[i];
            var pathName = Utils.safeTrim(pathItem.name) || ("Path_" + i);
            
            if (Utils.matchText(searchText, "", pathName)) {
                try {
                    this.foundElements.push({
                        name: pathName,
                        type: "PathItem",
                        position: pathItem.position,
                        bounds: pathItem.geometricBounds,
                        container: containerName,
                        opacity: pathItem.opacity || 100,
                        visible: !pathItem.hidden,
                        locked: pathItem.locked || false,
                        filled: pathItem.filled,
                        stroked: pathItem.stroked
                    });
                    
                    Utils.log("PathItem 발견: " + pathName + " 위치: " + Utils.formatPosition(pathItem.position));
                } catch (e) {
                    Utils.log("PathItem 정보 추출 실패: " + pathName + " - " + e.message);
                }
            }
        }
    },
    
    // ✅ 검색 결과 초기화
    clearResults: function() {
        this.foundElements = [];
    }
};

// ===== RESULT FORMATTER =====
var ResultFormatter = {
    generateReport: function(elements, searchText) {
        if (elements.length === 0) {
            return "❌ 검색 결과 없음\n\n검색어: '" + searchText + "'에 해당하는 요소를 찾을 수 없습니다.";
        }
        
        var report = "🔍 검색 결과: '" + searchText + "'\n";
        report += "발견된 요소: " + elements.length + "개\n";
        report += "=" * 50 + "\n\n";
        
        for (var i = 0; i < elements.length; i++) {
            var element = elements[i];
            
            report += "📍 " + (i + 1) + ". " + element.name + " (" + element.type + ")\n";
            report += "   위치: " + Utils.formatPosition(element.position) + "\n";
            report += "   경계: " + Utils.formatBounds(element.bounds) + "\n";
            report += "   컨테이너: " + element.container + "\n";
            report += "   투명도: " + element.opacity + "%\n";
            report += "   표시: " + (element.visible ? "보임" : "숨김") + "\n";
            report += "   잠금: " + (element.locked ? "잠김" : "해제") + "\n";
            
            // 타입별 추가 정보
            if (element.type === "TextFrame" && element.content) {
                report += "   내용: '" + element.content + "'\n";
            } else if (element.type === "PlacedItem" && element.fileName) {
                report += "   파일: " + element.fileName + "\n";
            } else if (element.type === "PathItem") {
                report += "   채우기: " + (element.filled ? "있음" : "없음") + "\n";
                report += "   선: " + (element.stroked ? "있음" : "없음") + "\n";
            }
            
            report += "\n";
        }
        
        return report;
    },
    
    generatePositionList: function(elements) {
        var positions = "📐 위치 요약:\n\n";
        
        for (var i = 0; i < elements.length; i++) {
            var element = elements[i];
            positions += element.name + " (" + element.type + "): " + Utils.formatPosition(element.position) + "\n";
        }
        
        return positions;
    }
};

// ===== MAIN CONTROLLER =====
var ElementPositionFinder = {
    run: function() {
        try {
            Utils.log("요소 위치 검색 시작");
            
            // 1. 문서 확인
            if (!app.activeDocument) {
                alert("❌ 활성 문서가 없습니다. 일러스트 파일을 열어주세요.");
                return;
            }
            
            var doc = app.activeDocument;
            Utils.log("문서명: " + doc.name);
            
            // 2. 검색어 입력
            var searchText = prompt("🔍 검색할 요소명 또는 텍스트를 입력하세요:\n\n" +
                "• 요소명 검색: 'Product-1', '개별포장', 'HACCP' 등\n" +
                "• 텍스트 내용 검색: '빠른 배송', '냉동보관' 등\n" +
                "• 부분 검색 지원: '개별' 입력시 '개별포장' 검색됨", "");
            
            if (!searchText || searchText === null) {
                Utils.log("검색 취소됨");
                return;
            }
            
            searchText = Utils.safeTrim(searchText);
            if (searchText.length === 0) {
                alert("❌ 검색어를 입력해주세요.");
                return;
            }
            
            Utils.log("검색어: '" + searchText + "'");
            
            // 3. 검색 실행
            ElementFinder.clearResults();
            
            // 모든 레이어에서 검색
            for (var i = 0; i < doc.layers.length; i++) {
                var layer = doc.layers[i];
                if (!layer.locked && layer.visible) {
                    ElementFinder.searchAllElements(layer, searchText, "레이어: " + layer.name);
                }
            }
            
            // 4. 결과 출력
            var foundElements = ElementFinder.foundElements;
            var report = ResultFormatter.generateReport(foundElements, searchText);
            
            Utils.log("검색 완료. 발견된 요소: " + foundElements.length + "개");
            
            // 상세 보고서 표시
            alert(report);
            
            // 위치 요약 표시 (요소가 많지 않은 경우)
            if (foundElements.length > 0 && foundElements.length <= 10) {
                var positionList = ResultFormatter.generatePositionList(foundElements);
                if (confirm("위치 요약을 추가로 보시겠습니까?")) {
                    alert(positionList);
                }
            }
            
            // 로그 출력 여부 확인
            if (foundElements.length > 0) {
                if (confirm("상세 로그를 Info 패널에서 확인하시겠습니까?\n(Window > Info 또는 F8키로 확인 가능)")) {
                    for (var j = 0; j < foundElements.length; j++) {
                        var el = foundElements[j];
                        Utils.log("=== " + el.name + " (" + el.type + ") ===");
                        Utils.log("위치: " + Utils.formatPosition(el.position));
                        Utils.log("경계: " + Utils.formatBounds(el.bounds));
                        Utils.log("컨테이너: " + el.container);
                    }
                }
            }
            
        } catch (e) {
            var errorMsg = "❌ 검색 중 오류 발생: " + e.message;
            alert(errorMsg);
            Utils.log("오류: " + e.toString());
        }
    }
};

// ===== 실행 =====
ElementPositionFinder.run();
