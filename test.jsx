var BASE_FOLDER_NAME = "작업파일";
var IMG_FOLDER_NAME = "img";
var TEMPLATE_LAYER_NAME = "auto_layer";
var CLIP_GROUP_PREFIX = "Clip Group-img";
var MINI_IMG_PREFIX = "mini-img";

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

// ✅ 클립 그룹용 이미지 파일 찾기
function findClipImageFile(imgFolderPath, pageNumber, groupNumber) {
    var extensions = [".jpg", ".jpeg", ".png", ".PNG", ".JPG", ".JPEG"];
    
    for (var i = 0; i < extensions.length; i++) {
        var fileName = pageNumber + "-" + groupNumber + extensions[i];
        var fullPath = imgFolderPath + "/" + fileName;
        var testFile = new File(fullPath);
        
        if (testFile.exists) {
            return {
                success: true,
                path: fullPath,
                extension: extensions[i],
                fileName: fileName
            };
        }
    }
    
    return {
        success: false,
        testedExtensions: extensions
    };
}

// ✅ 미니 이미지용 파일 찾기 (새로 추가)
function findMiniImageFile(imgFolderPath, pageNumber, groupNumber) {
    var extensions = [".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG"];
    
    for (var i = 0; i < extensions.length; i++) {
        var fileName = pageNumber + "-" + groupNumber + "-min" + extensions[i];
        var fullPath = imgFolderPath + "/" + fileName;
        var testFile = new File(fullPath);
        
        if (testFile.exists) {
            return {
                success: true,
                path: fullPath,
                extension: extensions[i],
                fileName: fileName
            };
        }
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

// ✅ 미니 이미지 프레임 개수 찾기 (새로 추가)
function countMiniImageFrames(container, prefix) {
    var count = 0;
    var foundFrames = [];
    
    // 직접 PlacedItem 검색
    for (var i = 0; i < container.placedItems.length; i++) {
        var placedItem = container.placedItems[i];
        var frameName = placedItem.name;
        
        if (frameName.indexOf(prefix) === 0) {
            var numberPart = frameName.substring(prefix.length);
            if (numberPart && !isNaN(parseInt(numberPart))) {
                var frameNumber = parseInt(numberPart);
                if (!arrayContains(foundFrames, frameNumber)) {
                    foundFrames.push(frameNumber);
                    count++;
                }
            }
        }
    }
    
    // 중첩 그룹에서도 검색
    for (var j = 0; j < container.groupItems.length; j++) {
        var parentGroup = container.groupItems[j];
        var nestedResult = countMiniImageFrames(parentGroup, prefix);
        
        for (var k = 0; k < nestedResult.foundFrames.length; k++) {
            var nestedNumber = nestedResult.foundFrames[k];
            if (!arrayContains(foundFrames, nestedNumber)) {
                foundFrames.push(nestedNumber);
                count++;
            }
        }
    }
    
    return {
        count: count,
        foundFrames: sortNumbers(foundFrames)
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

// ✅ 미니 이미지 프레임 찾기 (새로 추가)
function findMiniImageFrame(frameNumber, prefix, container) {
    var targetName = prefix + frameNumber;
    
    // 직접 PlacedItem 검색
    for (var i = 0; i < container.placedItems.length; i++) {
        var placedItem = container.placedItems[i];
        if (placedItem.name === targetName) {
            return placedItem;
        }
    }
    
    // 중첩 그룹에서 검색
    for (var j = 0; j < container.groupItems.length; j++) {
        var parentGroup = container.groupItems[j];
        var found = findMiniImageFrame(frameNumber, prefix, parentGroup);
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

// ✅ 클립 그룹 + 미니 이미지 통합 처리 메인 함수
function linkAllProductImages() {
    try {
        var summary = "🖼️ 클립 그룹 + 미니 이미지 통합 연결 결과:\n\n";
        
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
        
        // ✅ 클립 그룹 개수 감지
        var clipGroupResult = countClipGroupsWithBrackets(templateLayer, CLIP_GROUP_PREFIX);
        var totalClipGroups = clipGroupResult.count;
        var foundClipNumbers = clipGroupResult.foundGroups;
        
        // ✅ 미니 이미지 프레임 개수 감지
        var miniFrameResult = countMiniImageFrames(templateLayer, MINI_IMG_PREFIX);
        var totalMiniFrames = miniFrameResult.count;
        var foundMiniNumbers = miniFrameResult.foundFrames;
        
        summary += "🔍 감지 결과:\n";
        summary += "• 클립 그룹: " + totalClipGroups + "개 (번호: " + arrayToString(foundClipNumbers, ", ") + ")\n";
        summary += "• 미니 이미지: " + totalMiniFrames + "개 (번호: " + arrayToString(foundMiniNumbers, ", ") + ")\n\n";
        
        if (totalClipGroups === 0 && totalMiniFrames === 0) {
            summary += "❌ 처리할 이미지 요소를 찾을 수 없습니다!\n\n";
            alert(summary);
            return;
        }
        
        var pageNumber = 31;
        var clipSuccessCount = 0;
        var miniSuccessCount = 0;
        var clipErrorCount = 0;
        var miniErrorCount = 0;
        
        // ✅ 1. 클립 그룹 처리
        if (totalClipGroups > 0) {
            summary += "=== 클립 그룹 이미지 연결 + 리사이징 처리 ===\n";
            
            for (var i = 0; i < foundClipNumbers.length; i++) {
                var groupNumber = foundClipNumbers[i];
                var targetGroupName = "<" + CLIP_GROUP_PREFIX + groupNumber + ">";
                summary += "--- " + targetGroupName + " 처리 ---\n";
                
                var clipGroup = findClipGroupByNameWithBrackets(groupNumber, CLIP_GROUP_PREFIX, templateLayer);
                
                if (!clipGroup) {
                    summary += "❌ 클립 그룹을 찾을 수 없음\n\n";
                    clipErrorCount++;
                    continue;
                }
                
                var foundItems = findPlacedItemAndRectangle(clipGroup);
                
                if (!foundItems.placedItem || !foundItems.rectangle) {
                    summary += "❌ PlacedItem 또는 Rectangle을 찾을 수 없음\n\n";
                    clipErrorCount++;
                    continue;
                }
                
                var imageResult = findClipImageFile(imgFolderPath, pageNumber, groupNumber);
                
                if (!imageResult.success) {
                    summary += "❌ 클립 이미지 파일 없음: " + pageNumber + "-" + groupNumber + ".*\n\n";
                    clipErrorCount++;
                    continue;
                }
                
                if (!relinkImage(foundItems.placedItem, imageResult.path)) {
                    summary += "❌ 클립 이미지 연결 실패\n\n";
                    clipErrorCount++;
                    continue;
                }
                
                var resizeResult = resizeImageToFitInsideRectangle(foundItems.placedItem, foundItems.rectangle);
                
                if (resizeResult.success) {
                    summary += "✅ 클립 이미지 연결 + 리사이징 성공: " + imageResult.fileName + " (스케일: " + Math.round(resizeResult.scale) + "%)\n\n";
                    clipSuccessCount++;
                } else {
                    summary += "❌ 클립 이미지 리사이징 실패\n\n";
                    clipErrorCount++;
                }
            }
        }
        
        // ✅ 2. 미니 이미지 처리 (새로 추가)
        if (totalMiniFrames > 0) {
            summary += "=== 미니 이미지 연결 처리 ===\n";
            
            for (var j = 0; j < foundMiniNumbers.length; j++) {
                var frameNumber = foundMiniNumbers[j];
                var targetFrameName = MINI_IMG_PREFIX + frameNumber;
                summary += "--- " + targetFrameName + " 처리 ---\n";
                
                var miniFrame = findMiniImageFrame(frameNumber, MINI_IMG_PREFIX, templateLayer);
                
                if (!miniFrame) {
                    summary += "❌ 미니 이미지 프레임을 찾을 수 없음: " + targetFrameName + "\n\n";
                    miniErrorCount++;
                    continue;
                }
                
                summary += "✅ 미니 프레임 발견: " + miniFrame.name + "\n";
                
                var miniImageResult = findMiniImageFile(imgFolderPath, pageNumber, frameNumber);
                
                if (!miniImageResult.success) {
                    summary += "❌ 미니 이미지 파일 없음. 시도한 파일명:\n";
                    for (var e = 0; e < miniImageResult.testedExtensions.length; e++) {
                        summary += "  • " + pageNumber + "-" + frameNumber + "-min" + miniImageResult.testedExtensions[e] + "\n";
                    }
                    summary += "\n";
                    miniErrorCount++;
                    continue;
                }
                
                summary += "✅ 미니 이미지 파일 발견: " + miniImageResult.fileName + "\n";
                
                if (relinkImage(miniFrame, miniImageResult.path)) {
                    summary += "✅ 미니 이미지 연결 성공!\n\n";
                    miniSuccessCount++;
                } else {
                    summary += "❌ 미니 이미지 연결 실패\n\n";
                    miniErrorCount++;
                }
            }
        }
        
        // ✅ 3. 최종 결과
        summary += "=== 최종 결과 ===\n";
        summary += "클립 그룹: 성공 " + clipSuccessCount + "개, 실패 " + clipErrorCount + "개\n";
        summary += "미니 이미지: 성공 " + miniSuccessCount + "개, 실패 " + miniErrorCount + "개\n";
        summary += "전체: 성공 " + (clipSuccessCount + miniSuccessCount) + "개, 실패 " + (clipErrorCount + miniErrorCount) + "개\n\n";
        
        if ((clipSuccessCount + miniSuccessCount) > 0) {
            summary += "🎉 총 " + (clipSuccessCount + miniSuccessCount) + "개 이미지가 성공적으로 연결되었습니다!\n\n";
        }
        
        summary += "🔧 처리된 기능:\n";
        summary += "✅ 클립 그룹: 자동 연결 + 비율 유지 리사이징\n";
        summary += "✅ 미니 이미지: 자동 연결 (리사이징 없음)\n";
        summary += "✅ 다중 확장자 지원 (jpg, png 등)\n";
        summary += "✅ 자동 개수 감지 및 중첩 그룹 지원\n";
        summary += "✅ 상대 경로 + ExtendScript 완전 호환\n\n";
        
        summary += "📁 파일명 패턴:\n";
        summary += "• 클립 그룹: " + pageNumber + "-[번호].jpg/png\n";
        summary += "• 미니 이미지: " + pageNumber + "-[번호]-min.jpg/png";
        
        alert(summary);
        
    } catch (e) {
        alert("❌ 스크립트 실행 오류: " + e.message + "\n라인: " + (e.line || '알 수 없음'));
    }
}

// ✅ 클립 그룹 + 미니 이미지 통합 자동 연결 실행
linkAllProductImages();
