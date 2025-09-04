// =====================================
// 오토단가 자동화 스크립트 - ExtendScript 완전 호환
// CSV 헤더 기반 동적 처리 + PlacedItem 지원
// =====================================

// ===== CONFIG 설정 =====
var CONFIG = {
    VER: "v2.0",
    FOLDERS: {
        BASE_FOLDER_NAME: "작업파일",
        IMG_FOLDER_NAME: "img",
        TEMPLATE_LAYER_NAME: "auto_layer"
    },
    PATTERNS: {
        PRODUCT_GROUP_PREFIX: "Product-",
        MINI_IMG_NAME: "mini-img",
        CLIP_GROUP_NAME: "<Clip Group-img>"
    },
    // ✅ 새로 추가된 분류 기준
    PRODUCT_TAGS: ["자연해동", "D-7발주", "개별포장"],
    CERTIFICATION_MARKS: ["HACCP", "유기가공식품", "전통식품", "품질인증",
        "무항생제", "무농약가공식품", "동물복지"],
    SPECIAL_ITEMS: ["NEW"],
    // ✅ 기본 필드에 새 필드들 추가
    BASIC_FIELDS: [
        "페이지", "순서", "타이틀", "상품명", "용량", "원재료", "보관방법",
        "설명란", "서브단가명", "서브단가", "메인단가명", "메인단가", "개당단가", "알러지"
    ],
    STORAGE_COLORS: {
        "냉동": {c: 100, m: 40, y: 0, k: 0},
        "냉장": {c: 100, m: 0, y: 100, k: 0},
        "상온": {c: 0, m: 0, y: 20, k: 80}
    },
    LAYOUT_OPTIONS: {
        CERTIFICATION: {
            LEFT_TO_RIGHT: "cert_ltr",
            RIGHT_TO_LEFT: "cert_rtl",
            TOP_TO_BOTTOM: "cert_ttb"
        },
        ADDITIONAL_TEXT: {
            LEFT_TO_RIGHT: "add_ltr",
            TOP_TO_BOTTOM: "add_ttb"
        }
    },
    PAGE_LAYOUTS: {
        "default": {
            certification: "cert_ltr",
            additionalText: "add_ttb"
        }
    },
    LAYOUT_SPACING: {
        CERTIFICATION: 5,
        ADDITIONAL_TEXT: 5
    },
    IMAGE_EXTENSIONS: [".jpg", ".jpeg", ".png", ".PNG", ".JPG", ".JPEG"],
    DEBUG_MODE: true,
    // ✅ 최대 상품 개수 확장
    MAX_PRODUCTS: 20
};

// ===== UTILS 모듈 (ExtendScript 호환) =====
var Utils = {
    safeTrim: function(str) {
        if (typeof str !== "string") return "";
        return str.replace(/^\s+|\s+$/g, '');
    },
    
    arrayContains: function(array, value) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === value) return true;
        }
        return false;
    },
    
    arrayToString: function(array, separator) {
        if (!separator) separator = ",";
        var result = "";
        for (var i = 0; i < array.length; i++) {
            result += array[i];
            if (i < array.length - 1) result += separator;
        }
        return result;
    },
    
    isBasicField: function(fieldName) {
        return Utils.arrayContains(CONFIG.BASIC_FIELDS, fieldName);
    },
    
    log: function(message) {
        if (CONFIG.DEBUG_MODE) {
            $.writeln("[DEBUG] " + message);
        }
    },
    
    safeLength: function(value) {
        if (typeof value !== "string") return 0;
        return value.length;
    },
    
    isProductTag: function(fieldName) {
        return Utils.arrayContains(CONFIG.PRODUCT_TAGS, fieldName);
    },
    
    isCertificationMark: function(fieldName) {
        return Utils.arrayContains(CONFIG.CERTIFICATION_MARKS, fieldName);
    }
};

// ===== ENHANCED UI DIALOG 모듈 =====
var EnhancedUI = {
    createMainDialog: function(availablePages, pageData, headers) {
        try {
            var dialog = new Window("dialog", "🚀 오토단가 동적 자동화 - 설정");
            dialog.orientation = "column";
            dialog.alignChildren = ["fill", "top"];
            dialog.spacing = 12;
            dialog.margins = 16;

            // 프로젝트 정보 섹션
            var infoPanel = dialog.add("panel", undefined, "📋 프로젝트 정보");
            infoPanel.orientation = "column";
            infoPanel.alignChildren = ["fill", "top"];
            infoPanel.margins = 12;

            var infoGroup = infoPanel.add("group");
            infoGroup.add("statictext", undefined, "✅ CSV 헤더 감지: " + (headers ? headers.length : 0) + "개");

            var headerText = infoPanel.add("edittext", undefined,
                headers ? Utils.arrayToString(headers, ", ") : "헤더 없음",
                {readonly: true});
            headerText.preferredSize = [450, 20];

            // ✅ 최대 처리 상품 개수 표시
            var maxProductsGroup = infoPanel.add("group");
            maxProductsGroup.add("statictext", undefined, "🔢 최대 처리 상품: " + CONFIG.MAX_PRODUCTS + "개");

            // 페이지 선택 섹션
            var pagePanel = dialog.add("panel", undefined, "📄 페이지 선택");
            pagePanel.orientation = "row";
            pagePanel.alignChildren = ["fill", "center"];
            pagePanel.margins = 12;

            pagePanel.add("statictext", undefined, "페이지 번호:");
            var pageInput = pagePanel.add("edittext", undefined,
                availablePages && availablePages.length > 0 ? availablePages[0].toString() : "1");
            pageInput.characters = 5;

            var availablePagesText = pagePanel.add("statictext", undefined,
                "사용 가능: " + (availablePages ? Utils.arrayToString(availablePages, ", ") : "없음"));

            // 배치 방식 선택 섹션
            var layoutPanel = dialog.add("panel", undefined, "🎨 배치 방식 선택");
            layoutPanel.orientation = "column";
            layoutPanel.alignChildren = ["fill", "top"];
            layoutPanel.margins = 12;

            // 인증정보 배치 선택
            var certGroup = layoutPanel.add("group");
            certGroup.orientation = "column";
            certGroup.alignChildren = ["left", "top"];
            certGroup.spacing = 5;
            certGroup.add("statictext", undefined, "🏆 인증정보 배치:");

            var certRadioGroup = certGroup.add("group");
            certRadioGroup.orientation = "row";
            certRadioGroup.spacing = 15;

            var certLTR = certRadioGroup.add("radiobutton", undefined, "좌→우");
            var certRTL = certRadioGroup.add("radiobutton", undefined, "우→좌");
            var certTTB = certRadioGroup.add("radiobutton", undefined, "상→하");
            certLTR.value = true;

            // 상품태그 배치 선택
            var tagGroup = layoutPanel.add("group");
            tagGroup.orientation = "column";
            tagGroup.alignChildren = ["left", "top"];
            tagGroup.spacing = 5;
            tagGroup.add("statictext", undefined, "🏷️ 상품태그 배치:");

            var tagRadioGroup = tagGroup.add("group");
            tagRadioGroup.orientation = "row";
            tagRadioGroup.spacing = 15;

            var tagLTR = tagRadioGroup.add("radiobutton", undefined, "좌→우");
            var tagTTB = tagRadioGroup.add("radiobutton", undefined, "상→하");
            tagTTB.value = true;

            // 미리보기 섹션
            var previewPanel = dialog.add("panel", undefined, "🔍 미리보기");
            previewPanel.orientation = "column";
            previewPanel.alignChildren = ["fill", "top"];
            previewPanel.margins = 12;

            var previewText = previewPanel.add("edittext", undefined, "",
                {readonly: true, multiline: true});
            previewText.preferredSize = [450, 150];

            // 미리보기 업데이트 함수
            var updatePreview = function() {
                try {
                    var inputPage = parseInt(pageInput.text);
                    if (isNaN(inputPage)) {
                        previewText.text = "❌ 올바른 페이지 번호를 입력하세요.";
                        return;
                    }

                    // 페이지 존재 확인
                    var pageExists = false;
                    if (availablePages && availablePages.length > 0) {
                        for (var i = 0; i < availablePages.length; i++) {
                            if (availablePages[i] === inputPage) {
                                pageExists = true;
                                break;
                            }
                        }
                    }

                    if (!pageExists) {
                        previewText.text = "❌ 존재하지 않는 페이지입니다. 사용 가능: " +
                            (availablePages ? Utils.arrayToString(availablePages, ", ") : "없음");
                        return;
                    }

                    var productDataArray = AutoDanggaAutomation.getTargetPageData(pageData, inputPage);
                    var preview = AutoDanggaAutomation.generatePreview(inputPage, productDataArray, headers);
                    previewText.text = preview;
                } catch (e) {
                    previewText.text = "미리보기 오류: " + e.message;
                    Utils.log("미리보기 업데이트 오류: " + e.message);
                }
            };

            // 이벤트 연결
            pageInput.onChanging = updatePreview;
            updatePreview();

            // 버튼 섹션
            var buttonGroup = dialog.add("group");
            buttonGroup.alignment = "center";
            buttonGroup.spacing = 15;

            var okButton = buttonGroup.add("button", undefined, "✅ 실행", {name: "ok"});
            var cancelButton = buttonGroup.add("button", undefined, "❌ 취소", {name: "cancel"});

            okButton.preferredSize = [80, 30];
            cancelButton.preferredSize = [80, 30];

            okButton.onClick = function() {
                try {
                    var inputPage = parseInt(pageInput.text);
                    if (isNaN(inputPage)) {
                        alert("❌ 올바른 페이지 번호를 입력하세요.");
                        return;
                    }

                    var pageExists = false;
                    if (availablePages && availablePages.length > 0) {
                        for (var i = 0; i < availablePages.length; i++) {
                            if (availablePages[i] === inputPage) {
                                pageExists = true;
                                break;
                            }
                        }
                    }

                    if (!pageExists) {
                        alert("❌ 존재하지 않는 페이지입니다.\n사용 가능한 페이지: " +
                            (availablePages ? Utils.arrayToString(availablePages, ", ") : "없음"));
                        return;
                    }

                    dialog.close(1);
                } catch (e) {
                    alert("버튼 클릭 오류: " + e.message);
                    Utils.log("OK 버튼 오류: " + e.message);
                }
            };

            cancelButton.onClick = function() {
                dialog.close(0);
            };

            // 결과 수집 및 반환
            if (dialog.show() == 1) {
                var selectedPage = parseInt(pageInput.text);

                var certLayout = null;
                if (certLTR.value) {
                    certLayout = CONFIG.LAYOUT_OPTIONS.CERTIFICATION.LEFT_TO_RIGHT;
                } else if (certRTL.value) {
                    certLayout = CONFIG.LAYOUT_OPTIONS.CERTIFICATION.RIGHT_TO_LEFT;
                } else if (certTTB.value) {
                    certLayout = CONFIG.LAYOUT_OPTIONS.CERTIFICATION.TOP_TO_BOTTOM;
                }

                var tagLayout = null;
                if (tagLTR.value) {
                    tagLayout = CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT;
                } else if (tagTTB.value) {
                    tagLayout = CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM;
                }

                return {
                    page: selectedPage,
                    layout: {
                        certification: certLayout,
                        additionalText: tagLayout
                    }
                };
            }

            return null;

        } catch (e) {
            alert("UI 생성 오류: " + e.message);
            Utils.log("UI 생성 오류: " + e.message);
            return null;
        }
    }
};

// ===== CSV PARSER 모듈 (수정: 헤더를 첫 번째 줄로) =====
// ===== CSV PARSER 모듈 (강화된 오류 처리) =====
var CSVParser = {
    parseCSVFile: function(csvFile) {
        try {
            // 파일 존재 확인
            if (!csvFile || !csvFile.exists) {
                throw new Error("CSV 파일이 존재하지 않습니다.");
            }

            csvFile.open("r");
            var content = csvFile.read();
            csvFile.close();
            
            if (!content || content.length === 0) {
                throw new Error("CSV 파일이 비어있습니다.");
            }

            var lines = this.safeSplitLines(content);
            
            // ✅ 파일 구조 검증 강화
            if (lines.length < 2) {
                throw new Error("❌ CSV 파일 구조 오류!\n\n" +
                    "파일에 최소 2줄이 있어야 합니다:\n" +
                    "1줄: 헤더 (페이지,순서,상품명,...)\n" +
                    "2줄: 데이터\n\n" +
                    "현재 파일 줄 수: " + lines.length);
            }

            // ✅ 헤더 파싱 및 검증 (첫 번째 줄)
            var headers = this.parseCSVLine(lines[0]);
            
            if (!headers || headers.length === 0) {
                throw new Error("❌ 헤더 파싱 실패!\n\n" +
                    "첫 번째 줄에서 헤더를 찾을 수 없습니다.\n" +
                    "CSV 파일의 첫 번째 줄이 헤더인지 확인하세요.\n\n" +
                    "첫 번째 줄 내용: " + lines[0]);
            }

            // ✅ 필수 헤더 확인
            var requiredHeaders = ["페이지", "순서", "상품명"];
            var missingHeaders = [];
            
            for (var r = 0; r < requiredHeaders.length; r++) {
                var found = false;
                for (var h = 0; h < headers.length; h++) {
                    if (Utils.safeTrim(headers[h]) === requiredHeaders[r]) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    missingHeaders.push(requiredHeaders[r]);
                }
            }
            
            if (missingHeaders.length > 0) {
                throw new Error("❌ 필수 헤더 누락!\n\n" +
                    "누락된 헤더: " + missingHeaders.join(", ") + "\n\n" +
                    "감지된 헤더: " + Utils.arrayToString(headers, ", ") + "\n\n" +
                    "CSV 파일의 첫 번째 줄에 필수 헤더가 있는지 확인하세요.");
            }

            Utils.log("✅ 헤더 파싱 성공: " + Utils.arrayToString(headers, ", "));
            
            var allProductData = [];
            var errorRows = [];
            
            // ✅ 데이터 파싱 (두 번째 줄부터) - 오류 행 추적
            for (var i = 1; i < lines.length; i++) {
                var line = lines[i];
                if (typeof line === "string" && Utils.safeLength(Utils.safeTrim(line)) > 0) {
                    try {
                        var cells = this.parseCSVLine(line);
                        
                        if (cells.length < headers.length) {
                            errorRows.push({
                                row: i + 1,
                                expected: headers.length,
                                actual: cells.length,
                                content: line.substring(0, 100) + (line.length > 100 ? "..." : "")
                            });
                            continue;
                        }
                        
                        var productData = this.convertToProductData(cells, headers);
                        if (productData) {
                            allProductData.push(productData);
                        }
                    } catch (rowError) {
                        errorRows.push({
                            row: i + 1,
                            error: rowError.message,
                            content: line.substring(0, 100) + (line.length > 100 ? "..." : "")
                        });
                    }
                }
            }

            // ✅ 데이터 검증
            if (allProductData.length === 0) {
                var errorMsg = "❌ 처리 가능한 데이터가 없습니다!\n\n";
                
                if (errorRows.length > 0) {
                    errorMsg += "오류가 발생한 행들:\n";
                    for (var e = 0; e < Math.min(errorRows.length, 3); e++) {
                        var err = errorRows[e];
                        errorMsg += "• " + err.row + "행: ";
                        if (err.expected) {
                            errorMsg += "열 개수 부족 (필요:" + err.expected + ", 실제:" + err.actual + ")";
                        } else if (err.error) {
                            errorMsg += err.error;
                        }
                        errorMsg += "\n";
                    }
                    if (errorRows.length > 3) {
                        errorMsg += "• ... 외 " + (errorRows.length - 3) + "개 행\n";
                    }
                }
                
                errorMsg += "\n헤더 개수: " + headers.length + "\n";
                errorMsg += "헤더: " + Utils.arrayToString(headers, ", ");
                
                throw new Error(errorMsg);
            }

            return {
                data: allProductData,
                headers: headers
            };

        } catch (e) {
            // ✅ 사용자 친화적 오류 메시지
            var userMessage = "";
            
            if (e.message.indexOf("❌") === 0) {
                // 이미 포맷된 오류 메시지
                userMessage = e.message;
            } else {
                // 일반적인 오류
                userMessage = "❌ CSV 파일 읽기 오류!\n\n" +
                    "오류 내용: " + e.message + "\n\n" +
                    "해결 방법:\n" +
                    "1. CSV 파일의 첫 번째 줄이 헤더인지 확인\n" +
                    "2. 헤더에 '페이지', '순서', '상품명' 포함 확인\n" +
                    "3. 모든 데이터 행의 열 개수가 헤더와 일치하는지 확인";
            }
            
            alert(userMessage);
            Utils.log("CSV 파싱 상세 오류: " + e.toString());
            throw e;
        }
    },

    // 나머지 함수들은 동일...
    safeSplitLines: function(content) {
        var result = [];
        var current = "";
        var inQuotes = false;
        
        for (var i = 0; i < content.length; i++) {
            var currentChar = content.charAt(i);
            if (currentChar === '"') {
                inQuotes = !inQuotes;
                current += currentChar;
            } else if ((currentChar === "\n" || currentChar === "\r") && !inQuotes) {
                if (Utils.safeLength(Utils.safeTrim(current)) > 0) {
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
        
        if (Utils.safeLength(Utils.safeTrim(current)) > 0) {
            result.push(current);
        }
        
        return result;
    },

    parseCSVLine: function(line) {
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
                result.push(this.cleanCell(current));
                current = "";
            } else {
                current += currentChar;
            }
            i++;
        }
        
        result.push(this.cleanCell(current));
        return result;
    },

    cleanCell: function(value) {
        if (typeof value !== "string") return "";
        value = Utils.safeTrim(value);
        if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
            value = value.substring(1, value.length - 1);
            value = value.replace(/""/g, '"');
        }
        return value;
    },

    convertToProductData: function(cells, headers) {
        try {
            var data = {};
            for (var i = 0; i < headers.length && i < cells.length; i++) {
                var header = Utils.safeTrim(headers[i]);
                var value = Utils.safeTrim(cells[i] || '');
                
                if (!header) continue;

                if (header === "페이지" || header === "순서") {
                    data[header] = parseInt(value) || 0;
                } else if (Utils.isBasicField(header)) {
                    data[header] = value;
                } else {
                    if (value.toUpperCase() === 'Y' || value.toUpperCase() === 'N') {
                        data[header] = value.toUpperCase() === 'Y';
                    } else {
                        data[header] = value;
                    }
                }
                
                Utils.log("필드 변환: " + header + " = " + data[header]);
            }
            return data;
        } catch (e) {
            Utils.log("데이터 변환 오류: " + e.message);
            return null;
        }
    }
};


// ===== GROUP MANAGER 모듈 (확장) =====
var GroupManager = {
    findProductGroup: function(groupNumber, layer) {
        var targetName = CONFIG.PATTERNS.PRODUCT_GROUP_PREFIX + groupNumber;
        for (var i = 0; i < layer.groupItems.length; i++) {
            var group = layer.groupItems[i];
            if (group.name === targetName) {
                return group;
            }
        }
        
        for (var j = 0; j < layer.groupItems.length; j++) {
            var parentGroup = layer.groupItems[j];
            var found = this.findProductGroupInNested(groupNumber, parentGroup);
            if (found) return found;
        }
        
        return null;
    },

    findProductGroupInNested: function(groupNumber, container) {
        var targetName = CONFIG.PATTERNS.PRODUCT_GROUP_PREFIX + groupNumber;
        for (var i = 0; i < container.groupItems.length; i++) {
            var group = container.groupItems[i];
            if (group.name === targetName) {
                return group;
            }
        }
        
        for (var j = 0; j < container.groupItems.length; j++) {
            var found = this.findProductGroupInNested(groupNumber, container.groupItems[j]);
            if (found) return found;
        }
        
        return null;
    },

    findTextFrameInGroup: function(frameName, group) {
        for (var i = 0; i < group.textFrames.length; i++) {
            var textFrame = group.textFrames[i];
            if (textFrame.name === frameName) {
                return textFrame;
            }
        }
        
        for (var j = 0; j < group.groupItems.length; j++) {
            var subGroup = group.groupItems[j];
            var found = this.findTextFrameInGroup(frameName, subGroup);
            if (found) return found;
        }
        
        return null;
    },

    findClipGroupInProductGroup: function(productGroup) {
        Utils.log("클리핑 그룹 검색: " + CONFIG.PATTERNS.CLIP_GROUP_NAME);
        for (var i = 0; i < productGroup.groupItems.length; i++) {
            var subGroup = productGroup.groupItems[i];
            Utils.log("그룹 확인: " + subGroup.name + ", 클리핑: " + subGroup.clipped);
            if (subGroup.name === CONFIG.PATTERNS.CLIP_GROUP_NAME && subGroup.clipped) {
                Utils.log("클리핑 그룹 발견: " + subGroup.name);
                return subGroup;
            }
        }
        
        for (var j = 0; j < productGroup.groupItems.length; j++) {
            var parentGroup = productGroup.groupItems[j];
            var found = this.findClipGroupInProductGroup(parentGroup);
            if (found) return found;
        }
        
        return null;
    },

    findPlacedItemInGroup: function(group) {
        Utils.log("상품 이미지 검색: " + group.name);
        var clipGroup = this.findClipGroupInProductGroup(group);
        if (clipGroup) {
            for (var i = 0; i < clipGroup.placedItems.length; i++) {
                var placedItem = clipGroup.placedItems[i];
                Utils.log("클리핑 그룹 내 상품이미지 발견: " + placedItem.name);
                return placedItem;
            }
        }
        return null;
    },

    findMiniImageInGroup: function(group) {
        for (var i = 0; i < group.placedItems.length; i++) {
            var placedItem = group.placedItems[i];
            if (placedItem.name === CONFIG.PATTERNS.MINI_IMG_NAME) {
                return placedItem;
            }
        }
        
        for (var j = 0; j < group.groupItems.length; j++) {
            var subGroup = group.groupItems[j];
            var found = this.findMiniImageInGroup(subGroup);
            if (found) return found;
        }
        
        return null;
    },

    findRectangleInGroup: function(group) {
        var clipGroup = this.findClipGroupInProductGroup(group);
        if (clipGroup) {
            for (var i = 0; i < clipGroup.pathItems.length; i++) {
                return clipGroup.pathItems[i];
            }
        }
        
        for (var j = 0; j < group.pathItems.length; j++) {
            return group.pathItems[j];
        }
        
        return null;
    }
};

// ===== ELEMENT FINDER 모듈 =====
var ElementFinder = {
    findElementByName: function(name, container) {
        var patterns = [name, "Group-" + name, "<" + name + ">"];
        Utils.log("요소 검색: " + name);

        // 1. PlacedItem 검색
        for (var i = 0; i < container.placedItems.length; i++) {
            var placedItem = container.placedItems[i];
            for (var p = 0; p < patterns.length; p++) {
                if (placedItem.name === patterns[p]) {
                    Utils.log('PlacedItem 발견: ' + placedItem.name);
                    return {type: 'PlacedItem', item: placedItem};
                }
            }
        }

        // 2. GroupItem 검색
        for (var j = 0; j < container.groupItems.length; j++) {
            var group = container.groupItems[j];
            for (var q = 0; q < patterns.length; q++) {
                if (group.name === patterns[q]) {
                    Utils.log('GroupItem 발견: ' + group.name);
                    return {type: 'GroupItem', item: group};
                }
            }
        }

        // 3. 중첩 검색
        for (var k = 0; k < container.groupItems.length; k++) {
            var parentGroup = container.groupItems[k];
            var found = this.findElementByName(name, parentGroup);
            if (found) return found;
        }

        return null;
    }
};

// ===== DYNAMIC PROCESSOR 모듈 =====
var DynamicProcessor = {
    processDynamicFields: function(productGroup, productData, headers, layoutConfig) {
        Utils.log('=== 동적 필드 처리: ' + productGroup.name + ' ===');
        var certificationItems = [];
        var additionalItems = [];

        for (var i = 0; i < headers.length; i++) {
            var fieldName = Utils.safeTrim(headers[i]);
            if (Utils.isBasicField(fieldName) || !fieldName) continue;

            var value = productData[fieldName];
            var itemResult = ElementFinder.findElementByName(fieldName, productGroup);

            if (itemResult) {
                try {
                    itemResult.item.opacity = 0;
                    Utils.log('요소 숨김: ' + fieldName);
                } catch (e) {
                    Utils.log('요소 숨김 실패: ' + fieldName + ' - ' + e.message);
                }

                var shouldShow = false;
                if (typeof value === "boolean" && value) {
                    shouldShow = true;
                } else if (typeof value === "string" &&
                    (Utils.safeTrim(value).toUpperCase() === 'Y' ||
                     (Utils.safeTrim(value).length > 0 && Utils.safeTrim(value).toUpperCase() !== 'N'))) {
                    shouldShow = true;
                }

                if (shouldShow) {
                    if (fieldName === 'NEW') {
                        itemResult.item.opacity = 100;
                        Utils.log('NEW 즉시 표시');
                    } else {
                        if (Utils.arrayContains(CONFIG.PRODUCT_TAGS, fieldName)) {
                            additionalItems.push({
                                name: fieldName,
                                item: itemResult.item,
                                type: itemResult.type
                            });
                            Utils.log('상품 태그로 분류: ' + fieldName);
                        } else if (Utils.arrayContains(CONFIG.CERTIFICATION_MARKS, fieldName)) {
                            certificationItems.push({
                                name: fieldName,
                                item: itemResult.item,
                                type: itemResult.type
                            });
                            Utils.log('인증마크로 분류: ' + fieldName);
                        } else {
                            if (typeof value === "boolean") {
                                certificationItems.push({
                                    name: fieldName,
                                    item: itemResult.item,
                                    type: itemResult.type
                                });
                            } else {
                                additionalItems.push({
                                    name: fieldName,
                                    item: itemResult.item,
                                    type: itemResult.type
                                });
                            }
                        }
                    }
                }
            }
        }

        if (certificationItems.length > 0) {
            this.arrangeItems(certificationItems, layoutConfig.certification, "인증마크");
        }

        if (additionalItems.length > 0) {
            this.arrangeItems(additionalItems, layoutConfig.additionalText, "추가텍스트");
        }

        Utils.log('=== 동적 필드 처리 완료 ===');
    },

    arrangeItems: function(items, layoutType, category) {
        if (items.length === 0) return;
        
        var baseItem = items[0].item;
        var startX = baseItem.position[0];
        var startY = baseItem.position[1];
        var spacing = (category === "인증마크") ? CONFIG.LAYOUT_SPACING.CERTIFICATION : CONFIG.LAYOUT_SPACING.ADDITIONAL_TEXT;

        Utils.log(category + ' 배치: ' + layoutType);

        if (layoutType === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.LEFT_TO_RIGHT ||
            layoutType === CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT) {
            var currentX = startX;
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                item.item.position = [currentX, startY];
                item.item.opacity = 100;
                var itemWidth = this.getItemWidth(item.item);
                currentX += itemWidth + spacing;
                Utils.log(category + ' 좌→우: ' + item.name);
            }
        } else if (layoutType === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.RIGHT_TO_LEFT) {
            var totalWidth = 0;
            for (var k = 0; k < items.length; k++) {
                totalWidth += this.getItemWidth(items[k].item);
                if (k < items.length - 1) totalWidth += spacing;
            }

            var currentX = startX - totalWidth;
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                item.item.position = [currentX, startY];
                item.item.opacity = 100;
                var itemWidth = this.getItemWidth(item.item);
                currentX += itemWidth + spacing;
                Utils.log(category + ' 우→좌: ' + item.name);
            }
        } else if (layoutType === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.TOP_TO_BOTTOM ||
                   layoutType === CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM) {
            var currentY = startY;
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                item.item.position = [startX, currentY];
                item.item.opacity = 100;
                var itemHeight = this.getItemHeight(item.item);
                currentY -= (itemHeight + spacing);
                Utils.log(category + ' 상→하: ' + item.name);
            }
        }
    },

    getItemWidth: function(item) {
        try {
            var bounds = item.geometricBounds;
            return Math.abs(bounds[2] - bounds[0]);
        } catch (e) {
            return 30;
        }
    },

    getItemHeight: function(item) {
        try {
            var bounds = item.geometricBounds;
            return Math.abs(bounds[1] - bounds[3]);
        } catch (e) {
            return 25;
        }
    }
};

// ===== TEXT PROCESSOR 모듈 =====
var TextProcessor = {
    applyStorageColor: function(textFrame, storageMethod) {
        try {
            var storage = Utils.safeTrim(storageMethod);
            var colorData = CONFIG.STORAGE_COLORS[storage];
            var cmykColor = new CMYKColor();
            
            if (colorData) {
                cmykColor.cyan = colorData.c;
                cmykColor.magenta = colorData.m;
                cmykColor.yellow = colorData.y;
                cmykColor.black = colorData.k;
            } else {
                cmykColor.cyan = 0;
                cmykColor.magenta = 0;
                cmykColor.yellow = 0;
                cmykColor.black = 100;
            }

            textFrame.textRange.characterAttributes.fillColor = cmykColor;
            Utils.log('색상 적용: ' + storage);
        } catch (e) {
            Utils.log('색상 적용 오류: ' + e.message);
        }
    },

    updateBasicFields: function(productGroup, productData) {
        Utils.log('기본 필드 업데이트: ' + productGroup.name);
        
        for (var i = 0; i < CONFIG.BASIC_FIELDS.length; i++) {
            var fieldName = CONFIG.BASIC_FIELDS[i];
            var textFrame = GroupManager.findTextFrameInGroup(fieldName, productGroup);
            
            if (textFrame) {
                var value = productData[fieldName] || '';
                textFrame.contents = value;
                Utils.log(fieldName + ' 업데이트: ' + value);
                
                if (fieldName === '보관방법' && value) {
                    this.applyStorageColor(textFrame, value);
                }
            }
        }
    }
};

// ===== IMAGE LINKER 모듈 =====
var ImageLinker = {
    getSiblingImageFolderPath: function(imgFolderName) {
        var docFile = app.activeDocument.fullName;
        var workFolder = docFile.parent;
        var parentFolder = workFolder.parent;
        
        if (!parentFolder) {
            throw new Error('상위 폴더를 찾을 수 없습니다.');
        }

        var imgFolder = new Folder(parentFolder.fsName + "/" + imgFolderName);
        if (!imgFolder.exists) {
            throw new Error('이미지 폴더를 찾을 수 없습니다: "' + imgFolder.fsName + '"');
        }

        return imgFolder.fsName;
    },

    findImageFile: function(imgFolderPath, pageNumber, productNumber, isMini) {
        for (var i = 0; i < CONFIG.IMAGE_EXTENSIONS.length; i++) {
            var fileName;
            if (isMini) {
                fileName = pageNumber + "-" + productNumber + "-min" + CONFIG.IMAGE_EXTENSIONS[i];
            } else {
                fileName = pageNumber + "-" + productNumber + CONFIG.IMAGE_EXTENSIONS[i];
            }

            var fullPath = imgFolderPath + "/" + fileName;
            var testFile = new File(fullPath);
            if (testFile.exists) {
                return {
                    success: true,
                    path: fullPath,
                    fileName: fileName
                };
            }
        }
        return { success: false };
    },

    resizeImageToFitRectangle: function(placedItem, rectangle) {
        try {
            var rectBounds = rectangle.geometricBounds;
            var rectWidth = Math.abs(rectBounds[2] - rectBounds[0]);
            var rectHeight = Math.abs(rectBounds[1] - rectBounds[3]);
            var imageWidth = placedItem.width;
            var imageHeight = placedItem.height;
            var scaleX = (rectWidth / imageWidth) * 100;
            var scaleY = (rectHeight / imageHeight) * 100;
            var finalScale = Math.min(scaleX, scaleY);

            placedItem.resize(finalScale, finalScale);

            var rectCenterX = (rectBounds[0] + rectBounds[2]) / 2;
            var rectCenterY = (rectBounds[1] + rectBounds[3]) / 2;
            var newImageBounds = placedItem.geometricBounds;
            var imageCenterX = (newImageBounds[0] + newImageBounds[2]) / 2;
            var imageCenterY = (newImageBounds[1] + newImageBounds[3]) / 2;
            var offsetX = rectCenterX - imageCenterX;
            var offsetY = rectCenterY - imageCenterY;

            placedItem.translate(offsetX, offsetY);
            return { success: true, scale: finalScale };
        } catch (e) {
            return { success: false, error: e.message };
        }
    },

    relinkImage: function(placedItem, imagePath) {
        try {
            var imageFile = new File(imagePath);
            placedItem.file = imageFile;
            return true;
        } catch (e) {
            return false;
        }
    },

    processGroupImages: function(productGroup, pageNumber, productNumber, imgFolderPath) {
        var results = { mainSuccess: false, miniSuccess: false, details: [] };

        // 메인 이미지 처리
        var mainPlacedItem = GroupManager.findPlacedItemInGroup(productGroup);
        if (mainPlacedItem) {
            var mainImageResult = this.findImageFile(imgFolderPath, pageNumber, productNumber, false);
            if (mainImageResult.success) {
                if (this.relinkImage(mainPlacedItem, mainImageResult.path)) {
                    results.details.push("메인 이미지 연결: " + mainImageResult.fileName);
                    var rectangle = GroupManager.findRectangleInGroup(productGroup);
                    if (rectangle) {
                        var resizeResult = this.resizeImageToFitRectangle(mainPlacedItem, rectangle);
                        if (resizeResult.success) {
                            results.details.push("리사이징 성공 (스케일: " + Math.round(resizeResult.scale) + "%)");
                        }
                    }
                    results.mainSuccess = true;
                }
            } else {
                results.details.push("메인 이미지 파일 없음");
            }
        } else {
            results.details.push("메인 이미지 PlacedItem 없음");
        }

        // 미니 이미지 처리
        var miniPlacedItem = GroupManager.findMiniImageInGroup(productGroup);
        if (miniPlacedItem) {
            miniPlacedItem.opacity = 0;
            var miniImageResult = this.findImageFile(imgFolderPath, pageNumber, productNumber, true);
            if (miniImageResult.success) {
                if (this.relinkImage(miniPlacedItem, miniImageResult.path)) {
                    miniPlacedItem.opacity = 100;
                    results.details.push("미니 이미지 연결: " + miniImageResult.fileName);
                    results.miniSuccess = true;
                }
            } else {
                results.details.push("미니 이미지 파일 없음");
            }
        }

        return results;
    }
};

// ===== MAIN 모듈 (오토단가로 변경 및 20개 확장) =====
var AutoDanggaAutomation = {
    getLayerByName: function(layerName, doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        return null;
    },

    getTargetPageData: function(allData, targetPage) {
        var pageProducts = [];
        for (var p = 0; p < allData.length; p++) {
            if (allData[p].페이지 === targetPage) {
                pageProducts.push(allData[p]);
            }
        }
        
        pageProducts.sort(function(a, b) {
            return a.순서 - b.순서;
        });
        
        // ✅ 수정: 최대 20개까지 처리
        return pageProducts.slice(0, CONFIG.MAX_PRODUCTS);
    },

    updateProductGroup: function(productNumber, productData, layer, headers, layoutConfig) {
        Utils.log("=== 상품 그룹 " + productNumber + " 업데이트 시작 ===");
        var productGroup = GroupManager.findProductGroup(productNumber, layer);
        
        if (!productGroup) {
            Utils.log("상품 그룹을 찾을 수 없음: Product-" + productNumber);
            return false;
        }

        TextProcessor.updateBasicFields(productGroup, productData);
        DynamicProcessor.processDynamicFields(productGroup, productData, headers, layoutConfig);
        
        Utils.log("=== 상품 그룹 " + productNumber + " 업데이트 완료 ===");
        return true;
    },

    getUserInput: function(allData, headers) {
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
        return EnhancedUI.createMainDialog(pages, allData, headers);
    },

    generatePreview: function(pageNumber, productDataArray, headers) {
        var preview = "오토단가 동적 자동화 - 페이지 " + pageNumber + ":\n\n";
        
        for (var j = 0; j < productDataArray.length; j++) {
            var product = productDataArray[j];
            var activeFields = [];

            for (var i = 0; i < headers.length; i++) {
                var fieldName = Utils.safeTrim(headers[i]);
                if (Utils.isBasicField(fieldName) || !fieldName) continue;
                
                var value = product[fieldName];
                if ((typeof value === "boolean" && value) ||
                    (typeof value === "string" && Utils.safeTrim(value).length > 0 &&
                     Utils.safeTrim(value).toUpperCase() !== 'N')) {
                    activeFields.push(fieldName + ": " + value);
                }
            }

            preview += "Product-" + (j + 1) + ": " + (product.상품명 || '') + "\n";
            preview += " 용량: " + (product.용량 || '') + "\n";
            preview += " 보관: " + (product.보관방법 || '') + "\n";
            preview += " 활성 필드: " + (activeFields.length > 0 ? Utils.arrayToString(activeFields, ", ") : "없음") + "\n\n";
        }

        preview += "🔍 감지된 헤더: " + Utils.arrayToString(headers, ", ") + "\n";
        preview += "📊 최대 처리 가능: " + CONFIG.MAX_PRODUCTS + "개 상품\n\n";
        
        return preview;
    },

    getLayoutDescription: function(layoutConfig) {
        var certDesc = "";
        var addDesc = "";
        
        if (layoutConfig.certification === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.LEFT_TO_RIGHT) {
            certDesc = "인증마크 좌→우";
        } else if (layoutConfig.certification === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.RIGHT_TO_LEFT) {
            certDesc = "인증마크 우→좌";
        } else if (layoutConfig.certification === CONFIG.LAYOUT_OPTIONS.CERTIFICATION.TOP_TO_BOTTOM) {
            certDesc = "인증마크 상→하";
        }

        if (layoutConfig.additionalText === CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT) {
            addDesc = "추가텍스트 좌→우";
        } else if (layoutConfig.additionalText === CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM) {
            addDesc = "추가텍스트 상→하";
        }

        return certDesc + ", " + addDesc;
    },

    run: function() {
        try {
            Utils.log("오토단가 동적 자동화 시작");

            var doc = app.activeDocument;
            var templateLayer = this.getLayerByName(CONFIG.FOLDERS.TEMPLATE_LAYER_NAME, doc);
            if (!templateLayer) {
                alert('❌ 템플릿 레이어 "' + CONFIG.FOLDERS.TEMPLATE_LAYER_NAME + '"를 찾을 수 없습니다.');
                return;
            }

            var csvFile = File.openDialog("오토단가 동적 자동화 CSV 파일 선택", "*.csv");
            if (!csvFile) return;
/////////
            // ✅ 안전한 CSV 파싱
            var csvResult = null;
            try {
                csvResult = CSVParser.parseCSVFile(csvFile);
            } catch (csvError) {
                // CSV 파싱 오류 시 여기서 멈춤 (alert은 CSVParser에서 이미 표시)
                return;
            }
            
            // ✅ 결과 유효성 재검사
            if (!csvResult || !csvResult.data || !csvResult.headers) {
                alert("❌ CSV 파싱 결과가 올바르지 않습니다.");
                return;
            }
            
            if (csvResult.data.length === 0) {
                alert("❌ 처리할 상품 데이터가 없습니다!\n\nCSV 파일에 데이터 행이 있는지 확인하세요.");
                return;
            }

            // ✅ UI 생성 전 데이터 검증
            var userInput = null;
            try {
                userInput = this.getUserInput(csvResult.data, csvResult.headers);
            } catch (uiError) {
                alert("❌ UI 생성 오류: " + uiError.message + "\n\nCSV 데이터 구조를 확인하세요.");
                return;
            }
            
            if (!userInput) return;
///////////////////////
            var targetPage = userInput.page;
            var layoutConfig = userInput.layout;
            var productDataArray = this.getTargetPageData(csvResult.data, targetPage);

            var layoutDescription = this.getLayoutDescription(layoutConfig);
            if (!confirm("🚀 오토단가 동적 자동화 실행\n\n" +
                "📄 선택된 페이지: " + targetPage + "\n" +
                "🎨 배치 방식: " + layoutDescription + "\n" +
                "📝 처리할 상품: " + productDataArray.length + "개 (최대 " + CONFIG.MAX_PRODUCTS + "개)\n\n" +
                "위 설정으로 자동화를 실행하시겠습니까?")) {
                return;
            }

            // ✅ 수정: 20개까지 처리
            var textUpdateCount = 0;
            for (var k = 0; k < productDataArray.length; k++) {
                var productNumber = k + 1;
                if (this.updateProductGroup(productNumber, productDataArray[k], templateLayer, csvResult.headers, layoutConfig)) {
                    textUpdateCount++;
                }
            }

            try {
                var imgFolderPath = ImageLinker.getSiblingImageFolderPath(CONFIG.FOLDERS.IMG_FOLDER_NAME);
                var imageResults = [];
                
                // ✅ 수정: 20개까지 이미지 처리
                for (var m = 0; m < productDataArray.length; m++) {
                    var productNumber = m + 1;
                    var productGroup = GroupManager.findProductGroup(productNumber, templateLayer);
                    if (productGroup) {
                        var imageResult = ImageLinker.processGroupImages(productGroup, targetPage, productNumber, imgFolderPath);
                        imageResults.push("Product-" + productNumber + ": " + imageResult.details.join(", "));
                    }
                }

                var successMessage = "✅ 오토단가 동적 자동화 완료!\n\n";
                successMessage += "🔍 헤더 감지: " + csvResult.headers.length + "개\n";
                successMessage += "📝 텍스트 업데이트: " + textUpdateCount + "개 그룹\n";
                successMessage += "🖼️ 이미지 처리:\n" + imageResults.join("\n") + "\n\n";
                successMessage += "🎨 적용된 배치: " + layoutDescription;
                alert(successMessage);
                
            } catch (imgError) {
                var partialSuccess = "⚠️ 부분 완료: 텍스트 처리 성공, 이미지 처리 오류\n\n";
                partialSuccess += "📝 텍스트 업데이트: " + textUpdateCount + "개 그룹\n";
                partialSuccess += "❌ 이미지 오류: " + imgError.message;
                alert(partialSuccess);
            }

        } catch (e) {
            alert("❌ 오토단가 자동화 오류: " + e.message + "\n라인: " + (e.line || '알 수 없음'));
            Utils.log("상세 오류: " + e.toString());
        }
    }
};

// ===== 실행 =====
AutoDanggaAutomation.run();
