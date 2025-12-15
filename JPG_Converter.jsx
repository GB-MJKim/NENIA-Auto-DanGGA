// ========================================
// AI 파일 일괄 JPG 변환 스크립트
// Adobe Illustrator ExtendScript (JSX)
// ========================================

#target illustrator

// 메인 함수 실행
main();

function main() {
    // 일러스트레이터가 실행 중인지 확인
    if (app.documents.length === 0) {
        // 폴더 선택 다이얼로그 표시
        var sourceFolder = Folder.selectDialog("AI 파일이 있는 폴더를 선택하세요");
        
        if (sourceFolder === null) {
            alert("폴더 선택이 취소되었습니다.");
            return;
        }
        
        // 출력 폴더 선택
        var outputFolder = Folder.selectDialog("JPG 파일을 저장할 폴더를 선택하세요");
        
        if (outputFolder === null) {
            alert("출력 폴더 선택이 취소되었습니다.");
            return;
        }
        
        // 가로 픽셀 입력받기
        var widthInput = prompt("JPG 가로 크기를 입력하세요 (px):", "1920");
        
        if (widthInput === null) {
            alert("가로 크기 입력이 취소되었습니다.");
            return;
        }
        
        // 입력값을 숫자로 변환
        var targetWidth = parseInt(widthInput, 10);
        
        // 유효성 검사
        if (isNaN(targetWidth) || targetWidth <= 0) {
            alert("올바른 숫자를 입력하세요.");
            return;
        }
        
        // AI 파일 목록 가져오기
        var aiFiles = getAIFiles(sourceFolder);
        
        if (aiFiles.length === 0) {
            alert("선택한 폴더에 AI 파일이 없습니다.");
            return;
        }
        
        // 변환 진행 확인
        var confirmMsg = aiFiles.length + "개의 AI 파일을 JPG로 변환하시겠습니까?\n\n";
        confirmMsg += "가로 크기: " + targetWidth + "px\n";
        confirmMsg += "세로 크기: 비율 자동 계산";
        
        if (!confirm(confirmMsg)) {
            return;
        }
        
        // 변환 처리
        processFiles(aiFiles, outputFolder, targetWidth);
        
        alert("변환 완료!\n" + aiFiles.length + "개의 파일이 처리되었습니다.");
        
    } else {
        alert("열려있는 문서를 모두 닫고 다시 실행해주세요.");
    }
}

// AI 파일 목록 가져오기 함수
function getAIFiles(folder) {
    var fileList = [];
    var files = folder.getFiles();
    
    // 폴더 내 모든 파일 검사
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        
        // 파일인지 확인 (폴더 제외)
        if (file instanceof File) {
            var fileName = file.name.toLowerCase();
            
            // .ai 파일만 추가
            if (fileName.length > 3 && fileName.substring(fileName.length - 3) === ".ai") {
                fileList.push(file);
            }
        }
    }
    
    return fileList;
}

// 파일 일괄 처리 함수
function processFiles(files, outputFolder, targetWidth) {
    var successCount = 0;
    var errorCount = 0;
    
    // 각 파일 처리
    for (var i = 0; i < files.length; i++) {
        try {
            var file = files[i];
            
            // AI 파일 열기
            var doc = app.open(file);
            
            // JPG로 저장
            exportToJPG(doc, outputFolder, targetWidth);
            
            // 문서 닫기 (저장하지 않음)
            doc.close(SaveOptions.DONOTSAVECHANGES);
            
            successCount++;
            
        } catch (e) {
            errorCount++;
            // 에러 발생 시 계속 진행
        }
    }
    
    // 결과 리포트
    if (errorCount > 0) {
        alert("완료: " + successCount + "개\n실패: " + errorCount + "개");
    }
}

// JPG 내보내기 함수
function exportToJPG(doc, outputFolder, targetWidth) {
    // 아트보드 크기 가져오기
    var artboard = doc.artboards[0];
    var artboardRect = artboard.artboardRect;
    
    // 아트보드 크기 계산 (left, top, right, bottom)
    var artboardWidth = artboardRect[2] - artboardRect[0];
    var artboardHeight = artboardRect[1] - artboardRect[3];
    
    // 비율 계산하여 세로 크기 자동 결정
    var aspectRatio = artboardHeight / artboardWidth;
    var targetHeight = Math.round(targetWidth * aspectRatio);
    
    // 출력 파일 경로 생성
    var fileName = doc.name;
    
    // .ai 확장자 제거
    if (fileName.length > 3) {
        var lastDot = fileName.lastIndexOf(".");
        if (lastDot > 0) {
            fileName = fileName.substring(0, lastDot);
        }
    }
    
    // JPG 파일 경로
    var jpgFile = new File(outputFolder + "/" + fileName + ".jpg");
    
    // JPG 내보내기 옵션 설정
    var exportOptions = new ExportOptionsJPEG();
    exportOptions.antiAliasing = true;
    exportOptions.artBoardClipping = true;
    exportOptions.qualitySetting = 100; // 최고 품질 (0-100)
    exportOptions.horizontalScale = (targetWidth / artboardWidth) * 100;
    exportOptions.verticalScale = (targetHeight / artboardHeight) * 100;
    
    // JPG 내보내기 실행
    doc.exportFile(jpgFile, ExportType.JPEG, exportOptions);
}
