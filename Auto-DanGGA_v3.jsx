// 오토단가 자동화 v3.0 (ExtendScript ES3 완전 호환)
// - 정규식 변수 사용 금지, 문자열 처리로만 구현

// [1] 설정
var CONFIG = {
    VER: "v3.0",
    FOLDERS: {
        BASE: "작업파일",
        IMG: "img"
    },
    PATTERNS: {
        PRODUCT_GROUP: "Product-",
        MINI_IMG: "mini-img"
    },
    BASIC_FIELDS: [
        "페이지", "순서", "타이틀", "상품명", "용량", "원재료", "보관방법",
        "설명란", "서브단가명", "서브단가", "메인단가명", "메인단가", "개당단가", "알러지"
    ],
    STORAGE_COLORS: {
        "냉동": {c:0, m:0, y:0, k:100}, 
        "냉장": {c:0, m:0, y:0, k:100}, 
        "상온": {c:0, m:0, y:0, k:100}
    },
    IMAGE_EXTENSIONS: [".jpg", ".jpeg", ".png", ".PNG", ".JPG", ".JPEG"],
    MAX_PRODUCTS: 20,
    DEBUG: false
};

// [2] 유틸 함수 (ES3 호환)
var Utils = {
    // 문자열 앞뒤 공백 제거 (정규식 사용 안 함)
    safeTrim: function(str) { 
        if(typeof str !== "string") return ""; 
        var start = 0;
        var end = str.length - 1;
        // 앞쪽 공백 제거
        while(start < str.length && (str.charAt(start) === " " || str.charAt(start) === "\t" || str.charAt(start) === "\n" || str.charAt(start) === "\r")) {
            start++;
        }
        // 뒤쪽 공백 제거
        while(end >= start && (str.charAt(end) === " " || str.charAt(end) === "\t" || str.charAt(end) === "\n" || str.charAt(end) === "\r")) {
            end--;
        }
        return str.substring(start, end + 1);
    },
    
    arrayContains: function(arr, v) { 
        for(var i=0; i<arr.length; i++) {
            if(arr[i] === v) return true;
        } 
        return false; 
    },
    
    log: function(msg) { 
        if(CONFIG.DEBUG) $.writeln("[DEBUG] " + msg); 
    }
};

// [3] CSV 파서 (ES3 완전 호환)
var CSVParser = {
    parse: function(csvFile) {
        if(!csvFile || !csvFile.exists) throw new Error("CSV 파일이 존재하지 않습니다");
        
        csvFile.open("r");
        var content = csvFile.read();
        csvFile.close();
        
        if(!content || content.length === 0) throw new Error("CSV 파일이 비어있습니다");
        
        // 안전한 줄 분리
        var lines = this.safeSplitLines(content);
        
        if(lines.length < 2) {
            throw new Error("CSV 구조 오류: 최소 2줄 필요 (1줄:헤더, 2줄~:데이터)");
        }
        
        // 헤더 파싱
        var headers = this.parseCSVLine(lines[0]);
        if(!headers || headers.length === 0) {
            throw new Error("헤더 파싱 실패");
        }
        
        // 데이터 파싱
        var data = [];
        for(var i=1; i<lines.length; i++) {
            var cells = this.parseCSVLine(lines[i]);
            if(cells.length === headers.length) {
                var row = {};
                for(var c=0; c<headers.length; c++) {
                    var key = Utils.safeTrim(headers[c]);
                    var val = Utils.safeTrim(cells[c] || "");
                    row[key] = val;
                }
                data.push(row);
            }
        }
        
        return {headers: headers, data: data};
    },
    
    // 줄 분리 (빈 줄 제거, 따옴표 내 줄바꿈 처리)
    safeSplitLines: function(content) {
        var result = [];
        var current = "";
        var inQuotes = false;
        
        for(var i=0; content.length; i++) {
            var ch = content.charAt(i);
            
            if(ch === '"') {
                inQuotes = !inQuotes;
                current += ch;
            } else if((ch === "\n" || ch === "\r") && !inQuotes) {
                var trimmed = Utils.safeTrim(current);
                if(trimmed.length > 0) {
                    result.push(current);
                }
                current = "";
                if(ch === "\r" && i+1 < content.length && content.charAt(i+1) === "\n") {
                    i++;
                }
            } else {
                current += ch;
            }
        }
        
        var trimmed = Utils.safeTrim(current);
        if(trimmed.length > 0) {
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
        
        while(i < line.length) {
            var ch = line.charAt(i);
            
            if(ch === '"') {
                if(inQuotes && i+1 < line.length && line.charAt(i+1) === '"') {
                    current += '"';
                    i += 2;
                    continue;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if(ch === "," && !inQuotes) {
                result.push(this.cleanCell(current));
                current = "";
            } else {
                current += ch;
            }
            i++;
        }
        
        result.push(this.cleanCell(current));
        return result;
    },
    
    cleanCell: function(value) {
        if(typeof value !== "string") return "";
        value = Utils.safeTrim(value);
        if(value.charAt(0) === '"' && value.charAt(value.length-1) === '"') {
            value = value.substring(1, value.length-1);
            // "" -> " 변환
            var replaced = "";
            for(var i=0; i<value.length; i++) {
                if(value.charAt(i) === '"' && i+1 < value.length && value.charAt(i+1) === '"') {
                    replaced += '"';
                    i++;
                } else {
                    replaced += value.charAt(i);
                }
            }
            value = replaced;
        }
        return value;
    }
};

// [4] 그룹/레이어 & 텍스트 처리
var GroupManager = {
    getLayerByName: function(name, doc) {
        for(var i=0; i<doc.layers.length; i++) {
            if(doc.layers[i].name === name) return doc.layers[i];
        }
        return null;
    },
    
    findProductGroup: function(num, layer) {
        var target = CONFIG.PATTERNS.PRODUCT_GROUP + num;
        for(var i=0; i<layer.groupItems.length; i++) {
            if(layer.groupItems[i].name === target) return layer.groupItems[i];
        }
        return null;
    },
    
    findTextFrame: function(field, group) {
        for(var i=0; i<group.textFrames.length; i++) {
            if(group.textFrames[i].name === field) return group.textFrames[i];
        }
        for(var j=0; j<group.groupItems.length; j++) {
            var res = this.findTextFrame(field, group.groupItems[j]);
            if(res) return res;
        }
        return null;
    }
};

var TextProcessor = {
    updateBasicFields: function(productGroup, data) {
        for(var i=0; i<CONFIG.BASIC_FIELDS.length; i++) {
            var field = CONFIG.BASIC_FIELDS[i];
            var tf = GroupManager.findTextFrame(field, productGroup);
            if(!tf) continue;
            
            tf.contents = (data[field] || "");
            
            if(field === "보관방법") {
                var storageValue = data[field];
                var color = CONFIG.STORAGE_COLORS[storageValue];
                if(color) {
                    var cmyk = new CMYKColor();
                    cmyk.cyan = color.c;
                    cmyk.magenta = color.m;
                    cmyk.yellow = color.y;
                    cmyk.black = color.k;
                    tf.textRange.characterAttributes.fillColor = cmyk;
                }
            }
        }
    }
};

// [5] 이미지 처리
var ImageLinker = {
    findImageFile: function(imgFolder, pageNum, prodNum, isMini) {
        for(var i=0; i<CONFIG.IMAGE_EXTENSIONS.length; i++) {
            var ext = CONFIG.IMAGE_EXTENSIONS[i];
            var fn = pageNum + "-" + prodNum + (isMini ? "-min" : "") + ext;
            var full = imgFolder + "/" + fn;
            var file = new File(full);
            if(file.exists) return {success: true, path: full, fileName: fn};
        }
        return {success: false};
    },
    
    relinkImage: function(item, path) {
        try {
            item.file = new File(path); 
            return true;
        } catch(e) {
            return false;
        }
    },
    
    processImages: function(productGroup, pageNum, prodNum, imgFolderPath) {
        for(var i=0; i<productGroup.placedItems.length; i++) {
            var pi = productGroup.placedItems[i];
            if(pi.name !== CONFIG.PATTERNS.MINI_IMG) {
                var main = ImageLinker.findImageFile(imgFolderPath, pageNum, prodNum, false);
                if(main.success) ImageLinker.relinkImage(pi, main.path);
            }
        }
        
        for(var j=0; j<productGroup.placedItems.length; j++) {
            var mini = productGroup.placedItems[j];
            if(mini.name === CONFIG.PATTERNS.MINI_IMG) {
                var miniFile = ImageLinker.findImageFile(imgFolderPath, pageNum, prodNum, true);
                if(miniFile.success) ImageLinker.relinkImage(mini, miniFile.path);
            }
        }
    }
};

// [6] 메인 실행
function runAutoDangga() {
    try {
        var doc = app.activeDocument;
        var layer = GroupManager.getLayerByName("auto_layer", doc);
        
        if(!layer) {
            alert("auto_layer 레이어 없음");
            return;
        }
        
        var csvFile = File.openDialog("CSV 선택", "*.csv");
        if(!csvFile) return;
        
        var csv = CSVParser.parse(csvFile);
        
        if(!csv.data || csv.data.length === 0) {
            alert("데이터 없음");
            return;
        }
        
        var pageSet = {};
        for(var i=0; isv.data.length; i++) {
            var pg = csv.data[i]["페이지"];
            if(pg && !pageSet[pg]) pageSet[pg] = true;
        }
        
        var pages = [];
        for(var k in pageSet) pages.push(k);
        pages.sort(function(a,b) {return parseInt(a) - parseInt(b);});
        
        var dlg = new Window("dialog", "페이지 선택");
        dlg.add("statictext", undefined, "페이지:");
        var input = dlg.add("edittext", undefined, pages.length ? pages[0] : "1");
        input.characters = 5;
        
        if(dlg.show() != 1) return;
        
        var targetPage = parseInt(input.text);
        
        var prods = [];
        for(var i=0; isv.data.length; i++) {
            if(parseInt(csv.data[i]["페이지"]) === targetPage) {
                prods.push(csv.data[i]);
            }
        }
        
        prods.sort(function(a,b) {
            return parseInt(a["순서"]) - parseInt(b["순서"]);
        });
        
        if(prods.length === 0) {
            alert("데이터 없음");
            return;
        }
        
        var updateCount = 0;
        for(var i=0; i<prods.length && i<CONFIG.MAX_PRODUCTS; i++) {
            var g = GroupManager.findProductGroup(i+1, layer);
            if(!g) continue;
            
            TextProcessor.updateBasicFields(g, prods[i]);
            
            var imgFolderPath = doc.fullName.parent.parent.fsName + "/" + CONFIG.FOLDERS.IMG;
            ImageLinker.processImages(g, targetPage, i+1, imgFolderPath);
            
            updateCount++;
        }
        
        alert("완료! " + updateCount + "개 처리");
        
    } catch(e) {
        alert("오류: " + e.message);
    }
}

runAutoDangga();
