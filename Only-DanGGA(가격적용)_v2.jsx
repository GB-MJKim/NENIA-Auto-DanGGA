/*
============================================================================
Auto DanGGA - 일러스트레이터 단가 자동 업데이트 스크립트
페이지-순서 기반 매칭으로 메인단가명, 메인단가, 서브단가명, 서브단가, 개당단가 변경
============================================================================
*/

// ============================================================================
// 0. 유틸리티 모듈
// ============================================================================
var Utils = {
    // ES3 호환 trim 함수
    safeTrim: function(str) {
        if (typeof str !== "string") return "";
        return str.replace(/^\s+|\s+$/g, "");
    },
    
    // 문자열 포함 여부
    contains: function(str, searchStr) {
        return str.indexOf(searchStr) !== -1;
    },
    
    // 파일명 디코딩
    decodeFileName: function(fileName) {
        try {
            return decodeURI(fileName);
        } catch (e) {
            return fileName;
        }
    },
    
    // 파일명에서 페이지 번호 추출 (월+페이지 형식 지원)
    extractPageNumber: function(fileName) {
        // "03월_12p_만두_1.ai" -> "12"
        // "24p_상품.ai" -> "24"
        var match = fileName.match(/(?:^\d+월_)?(\d+)p_/);
        if (match && match[1]) {
            return match[1];
        }
        return null;
    },

    // 경로 정규화 (백슬래시 처리)
    normalizePath: function(path) {
        if (typeof path !== "string") return "";
        // 슬래시를 백슬래시로 통일 (Windows)
        return path.replace(/\//g, "\\");
    },
    
    // 경로 유효성 검사
    isValidPath: function(path) {
        if (typeof path !== "string" || path === "") return false;
        
        var folder = new Folder(path);
        return folder.exists;
    }
};

// ============================================================================
// 1. CSV 파서 모듈
// ============================================================================
var CSVParser = {
    // CSV 파일 파싱
    parseCSVFile: function(csvFile) {
        try {
            if (!csvFile || !csvFile.exists) {
                throw new Error("CSV 파일을 찾을 수 없습니다.");
            }
            
            csvFile.encoding = "UTF-8";
            csvFile.open("r");
            var content = csvFile.read();
            csvFile.close();
            
            if (!content || content.length === 0) {
                throw new Error("CSV 파일이 비어있습니다.");
            }
            
            // CSV 파싱
            var lines = this.safeSplitLines(content);
            if (lines.length < 2) {
                throw new Error("CSV 파일에 데이터가 없습니다!");
            }
            
            var headers = this.parseCSVLine(lines[0]);
            if (!headers || headers.length === 0) {
                throw new Error("헤더를 읽을 수 없습니다!");
            }
            
            var allProductData = [];
            var pageNumbers = {};
            
            // 각 라인 파싱
            for (var i = 1; i < lines.length; i++) {
                var line = lines[i];
                if (typeof line === "string" && Utils.safeTrim(line).length === 0) {
                    continue;
                }
                
                try {
                    var cells = this.parseCSVLine(line);
                    if (cells.length !== headers.length) {
                        continue;
                    }
                    
                    var productData = this.convertToProductData(cells, headers);
                    if (productData) {
                        allProductData.push(productData);
                        
                        // 페이지 번호 수집
                        var pageNum = productData["페이지"];
                        if (pageNum) {
                            var pageInt = parseInt(pageNum, 10);
                            pageNumbers[pageInt] = true;
                        }
                    }
                } catch (rowError) {
                    continue;
                }


            }

            
            if (allProductData.length === 0) {
                throw new Error("파싱된 데이터가 없습니다!");
            }
            
            return {
                data: allProductData,
                headers: headers,
                pageNumbers: pageNumbers
            };
            
        } catch (e) {
            alert("CSV 파일 읽기 오류!\n" + e.message);
            throw e;
        }
    },
    
    // 라인 분리
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
                if (Utils.safeTrim(current).length > 0) {
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
        
        if (Utils.safeTrim(current).length > 0) {
            result.push(current);
        }
        
        return result;
    },
    
    // CSV 라인 파싱
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
    
    // 셀 정리
    cleanCell: function(value) {
        if (typeof value !== "string") {
            return "";
        }
        
        value = Utils.safeTrim(value);
        
        if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
            value = value.substring(1, value.length - 1);
        }
        
        value = value.replace(/""/g, '"');
        return value;
    },
    
    // 데이터 변환
    convertToProductData: function(cells, headers) {
        try {
            var data = {};
            
            for (var i = 0; i < headers.length && i < cells.length; i++) {
                var header = Utils.safeTrim(headers[i]);
                var value = Utils.safeTrim(cells[i]);
                
                if (!header) continue;
                
                data[header] = value;
            }
            
            return data;
        } catch (e) {
            return null;
        }
    },
    
    // ✅ 수정: 페이지-순서 찾기 (정수 비교)
    findByPageAndOrder: function(data, pageNum, orderNum) {
        var targetPage = parseInt(pageNum, 10);
        var targetOrder = parseInt(orderNum, 10);
        
        for (var i = 0; i < data.length; i++) {
            var csvPage = parseInt(data[i]["페이지"], 10);
            var csvOrder = parseInt(data[i]["순서"], 10);
            
            if (csvPage === targetPage && csvOrder === targetOrder) {
                return data[i];
            }
        }
        return null;
    }
};

// ============================================================================
// 2. 그룹 매니저 모듈
// ============================================================================
var GroupManager = {
    // 그룹 내에서 텍스트 프레임 찾기
    findTextFrameInGroup: function(frameName, group) {
        // 직접 텍스트 프레임 검색
        for (var i = 0; i < group.textFrames.length; i++) {
            var textFrame = group.textFrames[i];
            if (textFrame.name === frameName) {
                return textFrame;
            }
        }
        
        // 하위 그룹 재귀 검색
        for (var j = 0; j < group.groupItems.length; j++) {
            var subGroup = group.groupItems[j];
            var found = this.findTextFrameInGroup(frameName, subGroup);
            if (found) {
                return found;
            }
        }
        
        return null;
    }
};

// ============================================================================
// 3. 일러스트레이터 문서 모듈
// ============================================================================
var IllustratorDoc = {
    // auto_layer 레이어 찾기
    findAutoLayer: function(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "auto_layer") {
                return doc.layers[i];
            }
        }
        return null;
    },
    
    // Product-N 그룹 찾기
    findProductGroup: function(layer, orderNum) {
        var groupName = "Product-" + orderNum;
        
        for (var i = 0; i < layer.groupItems.length; i++) {
            if (layer.groupItems[i].name === groupName) {
                return layer.groupItems[i];
            }
        }
        
        return null;
    }
};

// ============================================================================
// 4. 텍스트 처리 모듈
// ============================================================================
var TextProcessor = {
    // 단가 필드 업데이트
    updatePriceFields: function(productGroup, csvData) {
        var priceFields = ["메인단가명", "메인단가", "서브단가명", "서브단가", "개당단가"];
        var updateCount = 0;
        var notFoundFields = [];
        
        for (var i = 0; i < priceFields.length; i++) {
            var fieldName = priceFields[i];
            var textFrame = GroupManager.findTextFrameInGroup(fieldName, productGroup);
            
            if (!textFrame) {
                notFoundFields.push(fieldName);
                continue;
            }
            
            var value = csvData[fieldName];
            
            // 값이 있으면 업데이트 (빈 문자열도 허용)
            if (value !== undefined && value !== null) {
                textFrame.contents = value;
                updateCount++;
            }
        }
        
        return {
            updateCount: updateCount,
            notFoundFields: notFoundFields
        };
    },

    // 상품코드 업데이트
    updateProductCode: function(productGroup, csvData) {
        var codeFrame = GroupManager.findTextFrameInGroup("상품코드", productGroup);
        if (!codeFrame) {
            return false;  // 그룹에 상품코드 요소 없음
        }
        var codeValue = csvData["상품코드"];
        if (codeValue === undefined || codeValue === null) {
            return false;  // CSV에 상품코드 없음
        }
        codeFrame.contents = Utils.safeTrim(codeValue);
        return true;
    },
    
    // 상품명 가져오기
    getProductName: function(productGroup) {
        var productNameFrame = GroupManager.findTextFrameInGroup("상품명", productGroup);
        if (productNameFrame) {
            return Utils.safeTrim(productNameFrame.contents);
        }
        return "";
    }
};

// ============================================================================
// 폴더 선택 UI 모듈 (경로 직접 입력 + 찾아보기 버튼)
// ============================================================================
var FolderSelectorUI = {
    // 폴더 선택 다이얼로그 (경로 입력 + 버튼)
    selectFolder: function() {
        var dialog = new Window("dialog", "폴더 선택");
        
        var mainGroup = dialog.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "fill";
        mainGroup.spacing = 10;
        
        // 안내 텍스트
        mainGroup.add("statictext", undefined, "처리할 폴더를 선택하거나 경로를 직접 입력하세요:");
        
        // 경로 입력 그룹
        var pathGroup = mainGroup.add("group");
        pathGroup.orientation = "row";
        pathGroup.alignChildren = "center";
        pathGroup.spacing = 10;
        
        // 경로 입력창
        var pathInput = pathGroup.add("edittext", undefined, "");
        pathInput.preferredSize = [450, 25];
        pathInput.active = true;
        
        // 찾아보기 버튼
        var browseBtn = pathGroup.add("button", undefined, "찾아보기...");
        browseBtn.preferredSize = [100, 25];
        
        var selectedFolder = null;
        
        // 찾아보기 버튼 클릭
        browseBtn.onClick = function() {
            var folder = Folder.selectDialog("처리할 폴더를 선택하세요");
            if (folder) {
                selectedFolder = folder;
                pathInput.text = folder.fsName;
            }
        };
        
        // 안내 메시지
        var infoGroup = mainGroup.add("group");
        infoGroup.orientation = "column";
        infoGroup.alignChildren = "left";
        
        infoGroup.add("statictext", undefined, "예시:");
        var exampleText = infoGroup.add("statictext", undefined, 
            "C:\\Users\\사용자명\\Documents\\작업폴더");
        exampleText.graphics.foregroundColor = exampleText.graphics.newPen(
            exampleText.graphics.PenType.SOLID_COLOR, 
            [0.5, 0.5, 0.5], 
            1
        );
        
        // 확인/취소 버튼
        var buttonGroup = dialog.add("group");
        buttonGroup.add("button", undefined, "확인", {name: "ok"});
        buttonGroup.add("button", undefined, "취소", {name: "cancel"});
        
        // 확인 버튼 클릭 시
        if (dialog.show() === 1) {
            var inputPath = Utils.safeTrim(pathInput.text);
            
            // 경로가 입력되었으면 검증
            if (inputPath !== "") {
                var folder = new Folder(inputPath);
                if (folder.exists) {
                    return folder;
                } else {
                    alert("입력한 경로가 존재하지 않습니다:\n" + inputPath);
                    return null;
                }
            }
            
            // 찾아보기로 선택된 폴더 반환
            return selectedFolder;
        }
        
        return null;
    }
};

// ============================================================================
// 5. 파일 선택 모듈
// ============================================================================
var FileSelector = {
    // 폴더 선택
    selectFolder: function() {
        return FolderSelectorUI.selectFolder();
    },
    
    // 폴더 내 일러스트 파일 목록
    getIllustratorFiles: function(folder) {
        var files = [];
        var fileList = folder.getFiles();
        
        for (var i = 0; i < fileList.length; i++) {
            var fileItem = fileList[i];
            
             // .ai, .pdf 파일 모두 수집
            if (fileItem instanceof File && 
                (Utils.contains(fileItem.name, ".ai") || 
                 Utils.contains(fileItem.name, ".pdf"))) {
                files.push(fileItem);
            }
        }
        
        return files;
    },
    
    // 파일을 CSV 존재 여부로 분류
    // ✅ 수정: 파일을 CSV 존재 여부로 분류 (정수 비교)
    categorizeFiles: function(fileList, csvPageNumbers) {
        var matched = [];
        var unmatched = [];
        
        for (var i = 0; i < fileList.length; i++) {
            var fileItem = fileList[i];
            var pageNum = Utils.extractPageNumber(fileItem.name);
            
                       
            // ✅ 수정: 정수로 변환해서 비교
            if (pageNum) {
                var pageInt = parseInt(pageNum, 10);
                if (csvPageNumbers[pageInt]) {
                    matched.push(fileItem);
                } else {
                    unmatched.push(fileItem);
                }
            } else {
                unmatched.push(fileItem);
            }
        }
        
        return {
            matched: matched,
            unmatched: unmatched
        };
    }
};

// ============================================================================
// 6. UI 다이얼로그 모듈 (스크롤바 추가, 고정 크기)
// ============================================================================
var UIDialog = {
    // 파일 선택 다이얼로그 생성
    createFileSelectionDialog: function(matchedFiles, unmatchedFiles) {
        var dialog = new Window("dialog", "파일 선택");
        
        var mainGroup = dialog.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "fill";
        mainGroup.spacing = 10;

        // UIDialog.createFileSelectionDialog 안에서, 패널들 아래 부분에 옵션 영역 추가
        var optionPanel = mainGroup.add("panel", undefined, "옵션");
        optionPanel.alignChildren = "left";
        optionPanel.margins = 15;

        // 구분선 삭제 체크박스
        var deleteDividerCheckbox = optionPanel.add(
            "checkbox",
            undefined,
            "사업단 가격 적용(구분선, 서브단가, 메인단가명 삭제)"
        );

        // === CSV에 존재하는 파일 섹션 ===
        var matchedPanel = mainGroup.add("panel", undefined, "CSV에 존재하는 페이지");
        matchedPanel.alignChildren = "fill";
        matchedPanel.margins = 10;
        
        var matchedButtonGroup = matchedPanel.add("group");
        matchedButtonGroup.add("button", undefined, "전체선택");
        matchedButtonGroup.add("button", undefined, "전체해제");
        
        // 스크롤 가능한 리스트박스로 변경
        var matchedListBox = matchedPanel.add("listbox", undefined, [], {multiselect: true});
        matchedListBox.preferredSize = [500, 200];
        
        for (var i = 0; i < matchedFiles.length; i++) {
            var displayName = Utils.decodeFileName(matchedFiles[i].name);
            matchedListBox.add("item", displayName);
        }
        
        // 전체선택/해제 이벤트
        matchedButtonGroup.children[0].onClick = function() {
            for (var k = 0; k < matchedListBox.items.length; k++) {
                matchedListBox.items[k].selected = true;
            }
        };
        
        matchedButtonGroup.children[1].onClick = function() {
            for (var k = 0; k < matchedListBox.items.length; k++) {
                matchedListBox.items[k].selected = false;
            }
        };
        
        // === CSV에 없는 파일 섹션 ===
        var unmatchedPanel = mainGroup.add("panel", undefined, "기타 파일");
        unmatchedPanel.alignChildren = "fill";
        unmatchedPanel.margins = 10;
        
        var unmatchedButtonGroup = unmatchedPanel.add("group");
        unmatchedButtonGroup.add("button", undefined, "전체선택");
        unmatchedButtonGroup.add("button", undefined, "전체해제");
        
        // 스크롤 가능한 리스트박스로 변경
        var unmatchedListBox = unmatchedPanel.add("listbox", undefined, [], {multiselect: true});
        unmatchedListBox.preferredSize = [500, 150];
        
        for (var j = 0; j < unmatchedFiles.length; j++) {
            var displayName2 = Utils.decodeFileName(unmatchedFiles[j].name);
            unmatchedListBox.add("item", displayName2);
        }
        
        // 전체선택/해제 이벤트
        unmatchedButtonGroup.children[0].onClick = function() {
            for (var k = 0; k < unmatchedListBox.items.length; k++) {
                unmatchedListBox.items[k].selected = true;
            }
        };
        
        unmatchedButtonGroup.children[1].onClick = function() {
            for (var k = 0; k < unmatchedListBox.items.length; k++) {
                unmatchedListBox.items[k].selected = false;
            }
        };
        
        // 버튼 그룹
        var okCancelGroup = dialog.add("group");
        okCancelGroup.add("button", undefined, "처리", {name: "ok"});
        okCancelGroup.add("button", undefined, "취소", {name: "cancel"});
        
        return {
            dialog: dialog,
            matchedFiles: matchedFiles,
            unmatchedFiles: unmatchedFiles,
            matchedListBox: matchedListBox,
            unmatchedListBox: unmatchedListBox,
            deleteDividerCheckbox: deleteDividerCheckbox
        };
    },
    
    // 다이얼로그 표시 및 선택된 파일 반환
    showDialog: function(matchedFiles, unmatchedFiles) {
        var ui = this.createFileSelectionDialog(matchedFiles, unmatchedFiles);
        
        if (ui.dialog.show() === 1) {
            var selectedFiles = [];
            
            // 매칭된 파일 체크
            for (var i = 0; i < ui.matchedListBox.items.length; i++) {
                if (ui.matchedListBox.items[i].selected) {
                    selectedFiles.push(ui.matchedFiles[i]);
                }
            }
            
            // 매칭 안 된 파일 체크
            for (var j = 0; j < ui.unmatchedListBox.items.length; j++) {
                if (ui.unmatchedListBox.items[j].selected) {
                    selectedFiles.push(ui.unmatchedFiles[j]);
                }
            }
            
            return {
            files: selectedFiles,
            deleteDivider: ui.deleteDividerCheckbox.value  // true/false
            };
        }
        
        return { files: [], deleteDivider: false };
    }
};

// ============================================================================
// 7. 처리 모듈
// ============================================================================
var Processor = {
    // 일러스트 파일 처리
    processFile: function(filePath, csvData, deleteDivider) {
        var result = {
            success: false,
            fileName: Utils.decodeFileName(new File(filePath).name),
            errors: [],
            warnings: []
        };
        
        try {
            var doc = app.open(new File(filePath));
            
            // 파일명에서 페이지 번호 추출
            var pageNum = Utils.extractPageNumber(new File(filePath).name);
            if (!pageNum) {
                result.errors.push("파일명에서 페이지 번호를 추출할 수 없음");
                doc.close(SaveOptions.DONOTSAVECHANGES);
                return result;
            }
            
            // auto_layer 찾기
            var autoLayer = IllustratorDoc.findAutoLayer(doc);
            if (!autoLayer) {
                result.errors.push("auto_layer 레이어를 찾을 수 없음");
                doc.close(SaveOptions.DONOTSAVECHANGES);
                return result;
            }
            
            var totalUpdates = 0;
            
            // Product-1, Product-2, ... 순회
            for (var orderNum = 1; orderNum <= 20; orderNum++) {
                var orderStr = orderNum.toString();
                var productGroup = IllustratorDoc.findProductGroup(autoLayer, orderStr);
                
                if (!productGroup) {
                    continue;
                }
                
                // CSV에서 해당 페이지-순서 데이터 찾기
                var csvRow = CSVParser.findByPageAndOrder(csvData, pageNum, orderStr);
                
                if (!csvRow) {
                    continue;
                }
                
                // 상품명 비교 (더블체크)
                var aiProductName = TextProcessor.getProductName(productGroup);
                var csvProductName = csvRow["상품명"] || "";
                
                if (aiProductName !== csvProductName) {
                    result.warnings.push(
                        "Product-" + orderStr + " 상품명 불일치: AI[" + aiProductName + "] ≠ CSV[" + csvProductName + "]"
                    );
                }
                
                // 단가 필드 업데이트 (상품명이 다르더라도 업데이트)
                var updateResult = TextProcessor.updatePriceFields(productGroup, csvRow);
                totalUpdates += updateResult.updateCount;

                // 상품코드 업데이트
                TextProcessor.updateProductCode(productGroup, csvRow);  // 성공/실패는 경고로만 보고해도 됨

                 // 구분선/가격 요소 삭제 옵션
                if (deleteDivider) {
                    GroupCleaner.deleteItemsByName(productGroup, {
                        "구분선": true,
                        "서브단가": true,
                        "메인단가명": true
                    });
                }
                
                if (updateResult.notFoundFields.length > 0) {
                    result.errors.push(
                        "Product-" + orderStr + " 필드 없음: " + updateResult.notFoundFields.join(", ")
                    );
                }
            }
            
            if (totalUpdates > 0) {
                doc.save();
                result.success = true;
            } else {
                result.errors.push("업데이트된 필드 없음");
            }
            
            doc.close(SaveOptions.DONOTSAVECHANGES);
            
        } catch (e) {
            result.errors.push("처리 오류: " + e.message);
        }
        
        return result;
    },
    
    // 여러 파일 일괄 처리
    processFiles: function(fileList, csvData, deleteDivider) {
        var results = [];
        var successCount = 0;
        var failCount = 0;
        
        for (var i = 0; i < fileList.length; i++) {
            var result = this.processFile(fileList[i].absoluteURI, csvData, deleteDivider);
            results.push(result);
            
            if (result.success) {
                successCount++;
            } else {
                failCount++;
            }
        }
        
        return {
            results: results,
            successCount: successCount,
            failCount: failCount
        };
    }
};

// ============================================================================
// 8. 결과 보고 모듈
// ============================================================================
var ReportGenerator = {
    showReport: function(processResult) {
        var reportText = "=== 처리 결과 ===\n\n";
        reportText += "총 " + processResult.successCount + " 파일 수정 완료\n";
        reportText += processResult.failCount + " 파일 수정 실패\n\n";
        
        // 경고 (상품명 불일치)
        var hasWarnings = false;
        for (var i = 0; i < processResult.results.length; i++) {
            var result = processResult.results[i];
            if (result.warnings && result.warnings.length > 0) {
                if (!hasWarnings) {
                    reportText += "--- 경고 (상품명 불일치) ---\n\n";
                    hasWarnings = true;
                }
                reportText += "[" + result.fileName + "]\n";
                for (var w = 0; w < result.warnings.length; w++) {
                    reportText += "  ⚠ " + result.warnings[w] + "\n";
                }
                reportText += "\n";
            }
        }
        
        // 실패 상세
        if (processResult.failCount > 0) {
            reportText += "--- 실패 파일 상세 ---\n\n";
            
            for (var j = 0; j < processResult.results.length; j++) {
                var result2 = processResult.results[j];
                if (!result2.success) {
                    reportText += "[" + result2.fileName + "]\n";
                    for (var k = 0; k < result2.errors.length; k++) {
                        reportText += "  - " + result2.errors[k] + "\n";
                    }
                    reportText += "\n";
                }
            }
        }
        
        var dialog = new Window("dialog", "처리 결과");
        
        var resultGroup = dialog.add("group");
        resultGroup.orientation = "column";
        resultGroup.alignChildren = "fill";
        
        var resultText = resultGroup.add("edittext", undefined, reportText, 
            {multiline: true, scrolling: true});
        resultText.preferredSize = [600, 400];
        resultText.readonly = true;
        
        dialog.add("button", undefined, "확인", {name: "ok"});
        
        dialog.show();
    }
};


var GroupCleaner = {
    // 이름으로 텍스트프레임/패스/그룹 삭제 (재귀)
    deleteItemsByName: function(group, targetNames) {
        // 텍스트 프레임
        for (var i = group.textFrames.length - 1; i >= 0; i--) {
            var tf = group.textFrames[i];
            if (targetNames[tf.name]) {
                tf.remove();
            }
        }
        // 패스(구분선이 PathItem일 수도 있어서)
        for (var j = group.pathItems.length - 1; j >= 0; j--) {
            var p = group.pathItems[j];
            if (targetNames[p.name]) {
                p.remove();
            }
        }
        // 하위 그룹 재귀
        for (var k = group.groupItems.length - 1; k >= 0; k--) {
            this.deleteItemsByName(group.groupItems[k], targetNames);
        }
    }
};


// ============================================================================
// 9. 메인 실행 모듈
// ============================================================================
var MainScript = {
    run: function() {
        try {
            if (typeof app === "undefined") {
                alert("일러스트레이터가 실행되어야 합니다.");
                return;
            }
            
            // CSV 파일 선택
            var csvFile = File.openDialog("CSV 파일을 선택하세요", "CSV files:*.csv");
            if (!csvFile) {
                alert("CSV 파일이 선택되지 않았습니다.");
                return;
            }
            
            // CSV 데이터 로드
            var csvResult = CSVParser.parseCSVFile(csvFile);
            if (!csvResult || !csvResult.data) {
                return;
            }
            
            alert("CSV 파일 로드 완료: " + csvResult.data.length + " 개의 상품");
            
            // 폴더 선택
            var folder = FileSelector.selectFolder();
            if (!folder) {
                alert("폴더가 선택되지 않았습니다.");
                return;
            }
            
            // 일러스트 파일 목록
            var fileList = FileSelector.getIllustratorFiles(folder);
            if (fileList.length === 0) {
                alert("선택한 폴더에 .ai 파일이 없습니다.");
                return;
            }
            
            // 파일 분류 (CSV 존재 여부)
            var categorized = FileSelector.categorizeFiles(fileList, csvResult.pageNumbers);
            
            alert("발견된 파일: " + fileList.length + "개\n" +
                  "- CSV 매칭: " + categorized.matched.length + "개\n" +
                  "- 기타: " + categorized.unmatched.length + "개");
            
            // 파일 선택 다이얼로그
            var selection = UIDialog.showDialog(categorized.matched, categorized.unmatched);
            var selectedFiles = selection.files;
            var deleteDivider = selection.deleteDivider;

            if (selectedFiles.length === 0) {
                alert("선택된 파일이 없습니다.");
                return;
            }

            // 파일 처리 시 옵션 전달
            var processResult = Processor.processFiles(selectedFiles, csvResult.data, deleteDivider);
            
            // 결과 보고
            ReportGenerator.showReport(processResult);
            
        } catch (e) {
            alert("실행 오류: " + e.message);
        }
    }
};

// ============================================================================
// 10. 스크립트 실행
// ============================================================================
MainScript.run();
