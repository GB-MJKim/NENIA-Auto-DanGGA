var BASE_FOLDER_NAME = "작업파일";
var IMG_FOLDER_NAME = "img";
var TEMPLATE_LAYER_NAME = "auto_layer";
var CLIP_GROUP_PREFIX = "Clip Group-img";

function trimString(str) {
    if (typeof str !== "string") return "";
    return str.replace(/^\s+|\s+$/g, '');
}

function arrayContains(array, value) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === value) {
            return true;
        }
    }
    return false;
}

function sortNumbers(array) {
    for (var i = 0; i < array.length - 1; i++) {
        for (var j = 0; j < array.length - i - 1; j++) {
            if (array[j] > array[j + 1]) {
                var temp = array[j];
                array[j] = array[j + 1];
                array[j + 1] = temp;
            }
        }
    }
    return array;
}

function arrayToString(array, separator) {
    var result = "";
    for (var i = 0; i < array.length; i++) {
        result += array[i];
        if (i < array.length - 1) {
            result += separator;
        }
    }
    return result;
}

function getSiblingImageFolderPath(imgFolderName) {
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
}

function findImageFileWithMultipleExtensions(imgFolderPath, pageNumber, groupNumber) {
    var extensions = [".jpg", ".jpeg", ".png", ".PNG", ".JPG", ".JPEG"];
    var foundFile = null;
    var foundExtension = "";
    
    for (var i = 0; i < extensions.length; i++) {
        var fileName = pageNumber + "-" + groupNumber + extensions[i];
        var fullPath = imgFolderPath + "/" + fileName;
        var testFile = new File(fullPath);
        
        if (testFile.exists) {
            foundFile = fullPath;
            foundExtension = extensions[i];
            break;
        }
    }
    
    if (foundFile) {
        return {
            success: true,
            path: foundFile,
            extension: foundExtension,
            fileName: pageNumber + "-" + groupNumber + foundExtension
        };
    }
    
    return {
        success: false,
        testedExtensions: extensions
    };
}

function countClipGroupsWithBrackets(container, prefix) {
    var count = 0;
    var foundGroups = [];
    
    for (var i = 0; i < container.groupItems.length; i++) {
        var group = container.groupItems[i];
        var groupName = group.name;
        
        var expectedStart = "<" + prefix;
        var expectedEnd = ">";
        
        if (groupName.indexOf(expectedStart) === 0 && 
            groupName.charAt(groupName.length - 1) === expectedEnd) {
            
            var startPos = expectedStart.length;
            var endPos = groupName.length - 1;
            var numberPart = groupName.substring(startPos, endPos);
            
            if (numberPart && !isNaN(parseInt(numberPart))) {
                var groupNumber = parseInt(numberPart);
                if (!arrayContains(foundGroups, groupNumber)) {
                    foundGroups.push(groupNumber);
                    count++;
                }
            }
        }
    }
    
    for (var j = 0; j < container.groupItems.length; j++) {
        var parentGroup = container.groupItems[j];
        var nestedResult = countClipGroupsWithBrackets(parentGroup, prefix);
        
        for (var k = 0; k < nestedResult.foundGroups.length; k++) {
            var nestedNumber = nestedResult.foundGroups[k];
            if (!arrayContains(foundGroups, nestedNumber)) {
                foundGroups.push(nestedNumber);
                count++;
            }
        }
    }
    
    return {
        count: count,
        foundGroups: sortNumbers(foundGroups)
    };
}

function findClipGroupByNameWithBrackets(groupNumber, prefix, container) {
    var targetName = "<" + prefix + groupNumber + ">";
    
    for (var i = 0; i < container.groupItems.length; i++) {
        var group = container.groupItems[i];
        if (group.name === targetName) {
            return group;
        }
    }
    
    for (var j = 0; j < container.groupItems.length; j++) {
        var parentGroup = container.groupItems[j];
        var found = findClipGroupByNameWithBrackets(groupNumber, prefix, parentGroup);
        if (found) {
            return found;
        }
    }
    
    return null;
}

function findPlacedItemAndRectangle(group) {
    var placedItem = null;
    var rectangle = null;
    
    for (var i = 0; i < group.pageItems.length; i++) {
        var item = group.pageItems[i];
        
        if (item.typename === "PlacedItem") {
            placedItem = item;
        } else if (item.typename === "PathItem" && item.name.indexOf("Rectangle") >= 0) {
            rectangle = item;
        } else if (item.typename === "PathItem" && !rectangle) {
            rectangle = item;
        }
    }
    
    if (!placedItem || !rectangle) {
        for (var j = 0; j < group.groupItems.length; j++) {
            var subGroup = group.groupItems[j];
            var found = findPlacedItemAndRectangle(subGroup);
            
            if (found.placedItem && !placedItem) {
                placedItem = found.placedItem;
            }
            if (found.rectangle && !rectangle) {
                rectangle = found.rectangle;
            }
        }
    }
    
    return {
        placedItem: placedItem,
        rectangle: rectangle
    };
}

function resizeImageToFitInsideRectangle(placedItem, rectangle) {
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
        
        return {
            success: true,
            scale: finalScale,
            rectSize: rectWidth + " x " + rectHeight,
            imageSize: imageWidth + " x " + imageHeight,
            fitType: finalScale === scaleX ? "가로 맞춤" : "세로 맞춤"
        };
        
    } catch (e) {
        return {
            success: false,
            error: e.message
        };
    }
}

function relinkImage(placedItem, imagePath) {
    try {
        var imageFile = new File(imagePath);
        placedItem.file = imageFile;
        return true;
    } catch (e) {
        return false;
    }
}

function linkAndResizeProductImagesInside() {
    try {
        var summary = "🖼️ 사각형 안에 완전히 들어가는 이미지 연결 결과:\n\n";
        
        var doc = app.activeDocument;
        summary += "📄 문서: " + doc.name + "\n";
        
        var templateLayer = null;
        for (var l = 0; l < doc.layers.length; l++) {
            if (doc.layers[l].name === TEMPLATE_LAYER_NAME) {
                templateLayer = doc.layers[l];
                break;
            }
        }
        
        if (!templateLayer) {
            alert("❌ 템플릿 레이어 '" + TEMPLATE_LAYER_NAME + "'를 찾을 수 없습니다!");
            return;
        }
        
        var imgFolderPath = "";
        try {
            imgFolderPath = getSiblingImageFolderPath(IMG_FOLDER_NAME);
            summary += "✅ 이미지 폴더: " + imgFolderPath + "\n\n";
        } catch (e) {
            summary += "❌ " + e.message + "\n\n";
            alert(summary);
            return;
        }
        
        var clipGroupResult = countClipGroupsWithBrackets(templateLayer, CLIP_GROUP_PREFIX);
        var totalClipGroups = clipGroupResult.count;
        var foundGroupNumbers = clipGroupResult.foundGroups;
        
        summary += "🔍 클립 그룹 검색 결과:\n";
        summary += "• 패턴: <" + CLIP_GROUP_PREFIX + "[숫자]>\n";
        summary += "• 발견된 개수: " + totalClipGroups + "개\n";
        summary += "• 발견된 번호: " + arrayToString(foundGroupNumbers, ", ") + "\n\n";
        
        if (totalClipGroups === 0) {
            summary += "❌ 클립 그룹을 찾을 수 없습니다!\n\n";
            alert(summary);
            return;
        }
        
        var pageNumber = 31;
        var processedCount = 0;
        var errorCount = 0;
        
        summary += "=== 사각형 안에 완전히 들어가는 리사이징 처리 ===\n";
        
        for (var i = 0; i < foundGroupNumbers.length; i++) {
            var groupNumber = foundGroupNumbers[i];
            var targetGroupName = "<" + CLIP_GROUP_PREFIX + groupNumber + ">";
            summary += "--- " + targetGroupName + " 처리 ---\n";
            
            var clipGroup = findClipGroupByNameWithBrackets(groupNumber, CLIP_GROUP_PREFIX, templateLayer);
            
            if (!clipGroup) {
                summary += "❌ 클립 그룹을 찾을 수 없음: " + targetGroupName + "\n\n";
                errorCount++;
                continue;
            }
            
            summary += "✅ 클립 그룹 발견: " + clipGroup.name + "\n";
            
            var foundItems = findPlacedItemAndRectangle(clipGroup);
            
            if (!foundItems.placedItem) {
                summary += "❌ <Linked File> PlacedItem을 찾을 수 없음\n\n";
                errorCount++;
                continue;
            }
            
            if (!foundItems.rectangle) {
                summary += "❌ <Rectangle> PathItem을 찾을 수 없음\n\n";
                errorCount++;
                continue;
            }
            
            summary += "✅ PlacedItem 발견: " + (foundItems.placedItem.name || "unnamed") + "\n";
            summary += "✅ Rectangle 발견: " + (foundItems.rectangle.name || "unnamed") + "\n";
            
            var imageResult = findImageFileWithMultipleExtensions(imgFolderPath, pageNumber, groupNumber);
            
            if (!imageResult.success) {
                summary += "❌ 이미지 파일 없음\n\n";
                errorCount++;
                continue;
            }
            
            summary += "✅ 이미지 파일 발견: " + imageResult.fileName + "\n";
            
            if (!relinkImage(foundItems.placedItem, imageResult.path)) {
                summary += "❌ 이미지 연결 실패\n\n";
                errorCount++;
                continue;
            }
            
            summary += "✅ 이미지 연결 성공\n";
            
            var resizeResult = resizeImageToFitInsideRectangle(foundItems.placedItem, foundItems.rectangle);
            
            if (resizeResult.success) {
                summary += "✅ 사각형 안 맞춤 성공 (스케일: " + Math.round(resizeResult.scale) + "%, " + resizeResult.fitType + ")\n";
                summary += "   Rectangle: " + resizeResult.rectSize + " pt\n";
                summary += "   원본 이미지: " + resizeResult.imageSize + " pt\n\n";
                processedCount++;
            } else {
                summary += "❌ 리사이징 실패: " + resizeResult.error + "\n\n";
                errorCount++;
            }
        }
        
        summary += "=== 최종 결과 ===\n";
        summary += "감지된 클립 그룹: " + totalClipGroups + "개\n";
        summary += "성공적으로 처리: " + processedCount + "개\n";
        summary += "실패: " + errorCount + "개\n\n";
        
        if (processedCount > 0) {
            summary += "🎉 " + processedCount + "개 이미지가 성공적으로 연결 및 리사이징되었습니다!\n\n";
        }
        
        summary += "🔧 개선된 기능:\n";
        summary += "✅ 상대 경로 자동 검색\n";
        summary += "✅ 다중 확장자 지원 (jpg, png 등)\n";
        summary += "✅ 클립 그룹 개수 자동 감지\n";
        summary += "✅ 비율 유지하며 사각형 안에 완전히 들어가게 리사이징\n";
        summary += "✅ 넘침 방지 (가로/세로 중 작은 스케일 적용)\n";
        summary += "✅ 중앙 정렬 자동 적용\n";
        summary += "✅ ExtendScript 완전 호환";
        
        alert(summary);
        
    } catch (e) {
        alert("❌ 스크립트 실행 오류: " + e.message + "\n라인: " + (e.line || '알 수 없음'));
    }
}

linkAndResizeProductImagesInside();
