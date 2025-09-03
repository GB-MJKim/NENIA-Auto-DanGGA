// =====================================
// 네니아 자동화 스크립트 - ExtendScript 완전 호환
// CSV 헤더 기반 동적 처리 + PlacedItem 지원
// =====================================

// ===== CONFIG 설정 =====
var CONFIG = {
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
    
    TEXT_FRAME_NAMES: {
        PRODUCT_NAME: "상품명",
        CAPACITY: "용량",
        INGREDIENTS: "원재료",
        STORAGE: "보관방법"
    },
    
    BASIC_FIELDS: ["페이지", "순서", "타이틀", "상품명", "용량", "원재료", "보관방법"],
    
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
    DEBUG_MODE: true
};

// ===== UTILS 모듈 (ExtendScript 호환) =====
var Utils = {
    safeTrim: function(str) {
        if (typeof str !== "string") return "";
        return str.replace(/^\s+|\s+$/g, '');
    },

    // ✅ ExtendScript 호환 배열 함수들
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
    }
};

// ===== LAYOUT SELECTOR 모듈 =====
var LayoutSelector = {
    getUserLayoutChoice: function() {
        var options = "배치 방식을 선택하세요:\n\n";
        options += "=== 인증정보 + 추가텍스트 조합 ===\n";
        options += "1. 인증: 좌→우, 추가: 상→하\n";
        options += "2. 인증: 우→좌, 추가: 상→하\n";
        options += "3. 인증: 좌→우, 추가: 좌→우\n";
        options += "4. 인증: 우→좌, 추가: 좌→우\n";
        options += "5. 인증: 상→하, 추가: 좌→우\n";
        options += "6. 인증: 상→하, 추가: 상→하\n\n";
        options += "선택 (1-6):";

        var choice = prompt(options, "1");
        if (!choice || choice === null) {
            return this.getDefaultLayout();
        }

        if (!/^[1-6]$/.test(choice)) {
            alert("잘못된 선택입니다. 기본값(1번)을 사용합니다.");
            return this.getDefaultLayout();
        }

        return this.parseLayoutChoice(choice);
    },

    parseLayoutChoice: function(choice) {
        var layouts = {
            "1": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.LEFT_TO_RIGHT,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM
            },
            "2": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.RIGHT_TO_LEFT,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM
            },
            "3": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.LEFT_TO_RIGHT,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT
            },
            "4": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.RIGHT_TO_LEFT,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT
            },
            "5": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.TOP_TO_BOTTOM,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.LEFT_TO_RIGHT
            },
            "6": {
                certification: CONFIG.LAYOUT_OPTIONS.CERTIFICATION.TOP_TO_BOTTOM,
                additionalText: CONFIG.LAYOUT_OPTIONS.ADDITIONAL_TEXT.TOP_TO_BOTTOM
            }
        };
        return layouts[choice];
    },

    getDefaultLayout: function() {
        return CONFIG.PAGE_LAYOUTS["default"];
    }
};

// ===== CSV PARSER 모듈 (동적 헤더 파싱) =====
var CSVParser = {
    parseCSVFile: function(csvFile) {
        csvFile.open("r");
        var content = csvFile.read();
        csvFile.close();
        
        if (!content || content.length === 0) {
            throw new Error("CSV 파일이 비어있습니다.");
        }

        var lines = this.safeSplitLines(content);
        if (lines.length < 3) {
            throw new Error("CSV 파일에 데이터가 부족합니다.");
        }

        // ✅ 동적 헤더 파싱 (2번째 줄)
        var headers = this.parseCSVLine(lines[1]);
        Utils.log("감지된 헤더: " + Utils.arrayToString(headers, ", "));
        
        var allProductData = [];
        
        // 데이터 파싱 (3번째 줄부터)
        for (var i = 2; i < lines.length; i++) {
            var line = lines[i];
            if (typeof line === "string" && Utils.safeLength(Utils.safeTrim(line)) > 0) {
                var cells = this.parseCSVLine(line);
                if (cells.length >= headers.length) {
                    var productData = this.convertToProductData(cells, headers);
                    allProductData.push(productData);
                }
            }
        }
        
        return {
            data: allProductData,
            headers: headers
        };
    },

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

    // ✅ 동적 데이터 변환 (헤더 기반)
    convertToProductData: function(cells, headers) {
        var data = {};
        
        for (var i = 0; i < headers.length && i < cells.length; i++) {
            var header = Utils.safeTrim(headers[i]);
            var value = Utils.safeTrim(cells[i] || '');
            
            if (!header) continue;
            
            // 기본 필드 처리
            if (header === "페이지" || header === "순서") {
                data[header] = parseInt(value) || 0;
            } else if (Utils.isBasicField(header)) {
                data[header] = value;
            } else {
                // 동적 필드 (인증마크/추가텍스트)
                if (value.toUpperCase() === 'Y' || value.toUpperCase() === 'N') {
                    data[header] = value.toUpperCase() === 'Y';
                } else {
                    data[header] = value;
                }
                Utils.log("동적 필드: " + header + " = " + data[header]);
            }
        }
        
        return data;
    }
};

// ===== GROUP MANAGER 모듈 =====
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

        // ✅ 헤더 기반 동적 처리
        for (var i = 0; i < headers.length; i++) {
            var fieldName = Utils.safeTrim(headers[i]);
            
            // 기본 필드 제외
            if (Utils.isBasicField(fieldName) || !fieldName) continue;
            
            var value = productData[fieldName];
            var itemResult = ElementFinder.findElementByName(fieldName, productGroup);
            
            if (itemResult) {
                // 모든 요소 먼저 숨김
                try {
                    itemResult.item.opacity = 0;
                    Utils.log('요소 숨김: ' + fieldName);
                } catch (e) {
                    Utils.log('요소 숨김 실패: ' + fieldName + ' - ' + e.message);
                }
                
                // 활성화 조건 확인
                var shouldShow = false;
                if (typeof value === "boolean" && value) {
                    shouldShow = true;
                } else if (typeof value === "string" && 
                          (Utils.safeTrim(value).toUpperCase() === 'Y' || 
                           (Utils.safeTrim(value).length > 0 && Utils.safeTrim(value).toUpperCase() !== 'N'))) {
                    shouldShow = true;
                }
                
                if (shouldShow) {
                    // NEW는 즉시 표시
                    if (fieldName === 'NEW') {
                        itemResult.item.opacity = 100;
                        Utils.log('NEW 즉시 표시');
                    } else {
                        // 분류하여 저장
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

        // 배치 처리
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
            // 좌→우
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
            // 우→좌
            var currentX = startX;
            for (var i = items.length - 1; i >= 0; i--) {
                var item = items[i];
                item.item.position = [currentX, startY];
                item.item.opacity = 100;
                var itemWidth = this.getItemWidth(item.item);
                currentX += itemWidth + spacing;
                Utils.log(category + ' 우→좌: ' + item.name);
            }
        } else {
            // 상→하
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
        
        var productNameFrame = GroupManager.findTextFrameInGroup(CONFIG.TEXT_FRAME_NAMES.PRODUCT_NAME, productGroup);
        if (productNameFrame) {
            productNameFrame.contents = productData.상품명 || '';
            Utils.log('상품명 업데이트: ' + productData.상품명);
        }

        var capacityFrame = GroupManager.findTextFrameInGroup(CONFIG.TEXT_FRAME_NAMES.CAPACITY, productGroup);
        if (capacityFrame) {
            capacityFrame.contents = productData.용량 || '';
            Utils.log('용량 업데이트: ' + productData.용량);
        }

        var ingredientsFrame = GroupManager.findTextFrameInGroup(CONFIG.TEXT_FRAME_NAMES.INGREDIENTS, productGroup);
        if (ingredientsFrame) {
            ingredientsFrame.contents = productData.원재료 || '';
            Utils.log('원재료 업데이트');
        }

        var storageFrame = GroupManager.findTextFrameInGroup(CONFIG.TEXT_FRAME_NAMES.STORAGE, productGroup);
        if (storageFrame && productData.보관방법) {
            storageFrame.contents = productData.보관방법;
            this.applyStorageColor(storageFrame, productData.보관방법);
            Utils.log('보관방법 업데이트: ' + productData.보관방법);
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

// ===== MAIN 모듈 =====
var NeniaGroupAutomation = {
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
        return pageProducts.slice(0, 4);
    },

    updateProductGroup: function(productNumber, productData, layer, headers, layoutConfig) {
        Utils.log("=== 상품 그룹 " + productNumber + " 업데이트 시작 ===");
        var productGroup = GroupManager.findProductGroup(productNumber, layer);
        
        if (!productGroup) {
            Utils.log("상품 그룹을 찾을 수 없음: Product-" + productNumber);
            return false;
        }

        // 1. 기본 텍스트 필드 업데이트
        TextProcessor.updateBasicFields(productGroup, productData);
        
        // 2. 동적 필드 처리 (헤더 기반)
        DynamicProcessor.processDynamicFields(productGroup, productData, headers, layoutConfig);
        
        Utils.log("=== 상품 그룹 " + productNumber + " 업데이트 완료 ===");
        return true;
    },

    selectTargetPage: function(allData) {
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
        
        var pageOptions = "네니아 동적 자동화 - 사용 가능한 페이지:\n\n";
        for (var j = 0; j < pages.length; j++) {
            pageOptions += "페이지 " + pages[j] + ": " + pageCount[pages[j]] + "개 상품\n";
        }

        var selectedPage = prompt(pageOptions + "\n적용할 페이지 번호를 입력하세요:");
        if (selectedPage === null) return null;
        
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
    },

    generatePreview: function(pageNumber, productDataArray, headers) {
        var preview = "네니아 동적 자동화 - 페이지 " + pageNumber + ":\n\n";
        
        for (var j = 0; j < productDataArray.length; j++) {
            var product = productDataArray[j];
            var activeFields = [];

            // 동적 필드 수집
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

        preview += "🔍 감지된 헤더: " + Utils.arrayToString(headers, ", ") + "\n\n";
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
            Utils.log("네니아 동적 자동화 시작");

            // 1. 문서 및 레이어 확인
            var doc = app.activeDocument;
            var templateLayer = this.getLayerByName(CONFIG.FOLDERS.TEMPLATE_LAYER_NAME, doc);
            if (!templateLayer) {
                alert('❌ 템플릿 레이어 "' + CONFIG.FOLDERS.TEMPLATE_LAYER_NAME + '"를 찾을 수 없습니다.');
                return;
            }

            // 2. CSV 파일 동적 파싱
            var csvFile = File.openDialog("네니아 동적 자동화 CSV 파일 선택", "*.csv");
            if (!csvFile) return;
            
            var csvResult = CSVParser.parseCSVFile(csvFile);
            if (csvResult.data.length === 0) {
                alert("처리할 상품 데이터가 없습니다!");
                return;
            }

            // 3. 페이지 선택
            var targetPage = this.selectTargetPage(csvResult.data);
            if (targetPage === null) return;

            // 4. 레이아웃 선택
            var layoutConfig = LayoutSelector.getUserLayoutChoice();

            // 5. 해당 페이지 데이터 필터링
            var productDataArray = this.getTargetPageData(csvResult.data, targetPage);

            // 6. 미리보기 및 확인
            var preview = this.generatePreview(targetPage, productDataArray, csvResult.headers);
            var layoutDescription = this.getLayoutDescription(layoutConfig);
            
            if (!confirm(preview + "위 내용으로 동적 자동화를 실행하시겠습니까?\n\n" +
                "✅ CSV 헤더 자동 인식\n" +
                "✅ 동적 요소 처리\n" +
                "✅ 클리핑 그룹 정확 처리\n" +
                "✅ ExtendScript 완전 호환\n" +
                "✅ 선택된 배치: " + layoutDescription)) {
                return;
            }

            // 7. 처리 실행
            var textUpdateCount = 0;
            for (var k = 0; k < productDataArray.length; k++) {
                var productNumber = k + 1;
                if (this.updateProductGroup(productNumber, productDataArray[k], templateLayer, csvResult.headers, layoutConfig)) {
                    textUpdateCount++;
                }
            }

            // 8. 이미지 처리
            try {
                var imgFolderPath = ImageLinker.getSiblingImageFolderPath(CONFIG.FOLDERS.IMG_FOLDER_NAME);
                var imageResults = [];

                for (var m = 0; m < productDataArray.length; m++) {
                    var productNumber = m + 1;
                    var productGroup = GroupManager.findProductGroup(productNumber, templateLayer);
                    
                    if (productGroup) {
                        var imageResult = ImageLinker.processGroupImages(productGroup, targetPage, productNumber, imgFolderPath);
                        imageResults.push("Product-" + productNumber + ": " + imageResult.details.join(", "));
                    }
                }

                // 9. 완료 메시지
                var successMessage = "✅ 네니아 동적 자동화 완료!\n\n";
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
            alert("❌ 자동화 오류: " + e.message + "\n라인: " + (e.line || '알 수 없음'));
            Utils.log("상세 오류: " + e.toString());
        }
    }
};

// ===== 실행 =====
NeniaGroupAutomation.run();
