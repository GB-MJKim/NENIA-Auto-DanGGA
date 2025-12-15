#target illustrator

// 메인 함수 실행
main();

function main() {
    // 열린 문서가 없으면 종료
    if (!documents.length) {
        alert("열려 있는 문서가 없습니다.");
        return;
    }
    
    var doc = app.activeDocument;
    var docTitle = doc.name.replace(/\.[^\.]+$/, ''); // 확장자 제거
    
    // CSV 파일 저장 위치 설정 (데스크톱)
    var csvFile = File('~/Desktop/' + docTitle + '_GroupedFontInfo.csv');
    
    // UTF-8 BOM 추가 (한글 깨짐 방지)
    var BOM = '\uFEFF';
    
    // CSV 헤더
    var csvContent = BOM + 'TextFrame명,구간번호,시작위치,끝위치,길이,내용,폰트명,폰트크기,색상,X좌표,Y좌표\r\n';
    
    // 모든 텍스트 프레임 순회
    for (var i = 0; i < doc.textFrames.length; i++) {
        var textFrame = doc.textFrames[i];
        var frameName = textFrame.name || ('TextFrame_' + (i + 1));
        var posX = Math.round(textFrame.position[0]);
        var posY = Math.round(textFrame.position[1]);
        
        // 텍스트 프레임의 모든 문자 순회
        var characters = textFrame.textRange.characters;
        
        if (characters.length === 0) continue;
        
        // 그룹 정보를 저장할 배열
        var groups = [];
        var currentGroup = null;
        
        for (var j = 0; j < characters.length; j++) {
            var currentChar = characters[j];
            
            try {
                var charAttr = currentChar.characterAttributes;
                
                // 현재 글자의 속성
                var fontName = charAttr.textFont ? charAttr.textFont.name : '정보없음';
                var fontSize = charAttr.size ? charAttr.size.toString() : '정보없음';
                
                // 색상 정보
                var colorInfo = '';
                if (charAttr.fillColor) {
                    var fillColor = charAttr.fillColor;
                    if (fillColor.typename == 'RGBColor') {
                        colorInfo = 'R:' + Math.round(fillColor.red) + 
                                   ' G:' + Math.round(fillColor.green) + 
                                   ' B:' + Math.round(fillColor.blue);
                    } else if (fillColor.typename == 'CMYKColor') {
                        colorInfo = 'C:' + Math.round(fillColor.cyan) + 
                                   ' M:' + Math.round(fillColor.magenta) + 
                                   ' Y:' + Math.round(fillColor.yellow) + 
                                   ' K:' + Math.round(fillColor.black);
                    }
                } else {
                    colorInfo = '정보없음';
                }
                
                // 현재 그룹이 없거나, 폰트/크기/색상이 다르면 새 그룹 생성
                if (!currentGroup || 
                    currentGroup.fontName !== fontName || 
                    currentGroup.fontSize !== fontSize || 
                    currentGroup.colorInfo !== colorInfo) {
                    
                    // 이전 그룹이 있으면 저장
                    if (currentGroup) {
                        groups.push(currentGroup);
                    }
                    
                    // 새 그룹 시작
                    currentGroup = {
                        startPos: j,
                        endPos: j,
                        fontName: fontName,
                        fontSize: fontSize,
                        colorInfo: colorInfo,
                        content: currentChar.contents
                    };
                } else {
                    // 같은 그룹이면 내용과 끝 위치만 업데이트
                    currentGroup.endPos = j;
                    currentGroup.content += currentChar.contents;
                }
                
            } catch (e) {
                // 에러 발생 시 현재 그룹 종료하고 새로 시작
                if (currentGroup) {
                    groups.push(currentGroup);
                }
                currentGroup = {
                    startPos: j,
                    endPos: j,
                    fontName: '정보없음',
                    fontSize: '정보없음',
                    colorInfo: '정보없음',
                    content: '오류'
                };
            }
        }
        
        // 마지막 그룹 저장
        if (currentGroup) {
            groups.push(currentGroup);
        }
        
        // 그룹 정보를 CSV에 추가
        for (var k = 0; k < groups.length; k++) {
            var group = groups[k];
            var length = group.endPos - group.startPos + 1;
            var content = group.content.replace(/[\r\n]/g, '↵').substring(0, 100); // 100자로 제한
            
            csvContent += '"' + frameName + '",' + 
                         (k + 1) + ',' + 
                         (group.startPos + 1) + ',' + 
                         (group.endPos + 1) + ',' + 
                         length + ',"' + 
                         content + '","' + 
                         group.fontName + '","' + 
                         group.fontSize + '","' + 
                         group.colorInfo + '",' + 
                         posX + ',' + 
                         posY + '\r\n';
        }
    }
    
    // CSV 파일 UTF-8로 저장
    try {
        csvFile.encoding = 'UTF-8';
        csvFile.open('w');
        csvFile.write(csvContent);
        csvFile.close();
        
        alert('그룹화된 폰트 정보가 성공적으로 저장되었습니다.\n위치: ' + csvFile.fsName + 
              '\n총 ' + doc.textFrames.length + '개의 텍스트 프레임이 분석되었습니다.');
    } catch (e) {
        alert('파일 저장 중 오류가 발생했습니다.\n' + e.message);
    }
}
