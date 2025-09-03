// =====================================
// 빈 텍스트 & 투명도 0인 이미지 삭제 - 고급 버전
// 다양한 옵션 및 보호 기능 포함
// =====================================

function advancedDeleteEmptyItems() {
    try {
        var doc = app.activeDocument;
        
        // 옵션 다이얼로그 생성
        var dialog = new Window('dialog', '빈 항목 삭제 - 고급 옵션');
        dialog.orientation = 'column';
        dialog.alignChildren = 'left';
        
        // 대상 레이어 선택
        var layerGroup = dialog.add('group');
        layerGroup.add('statictext', undefined, '대상 레이어:');
        var layerRadio1 = layerGroup.add('radiobutton', undefined, '활성 레이어만');
        var layerRadio2 = layerGroup.add('radiobutton', undefined, 'auto_layer만');
        var layerRadio3 = layerGroup.add('radiobutton', undefined, '전체 레이어');
        layerRadio1.value = true;
        
        // 삭제 옵션
        var optionGroup = dialog.add('group');
        optionGroup.orientation = 'column';
        optionGroup.add('statictext', undefined, '삭제 옵션:');
        var deleteTextCheck = optionGroup.add('checkbox', undefined, '빈 텍스트 프레임 삭제');
        var deleteImageCheck = optionGroup.add('checkbox', undefined, '투명도 0인 이미지 삭제');
        var deleteInvisibleCheck = optionGroup.add('checkbox', undefined, '완전 투명한 모든 객체 삭제');
        deleteTextCheck.value = true;
        deleteImageCheck.value = true;
        
        // 보호 옵션
        var protectGroup = dialog.add('group');
        protectGroup.orientation = 'column';
        protectGroup.add('statictext', undefined, '보호 옵션:');
        var protectTemplateCheck = protectGroup.add('checkbox', undefined, '템플릿 요소 보호 (페이지, 순서1-4 등)');
        var protectNamedCheck = protectGroup.add('checkbox', undefined, '이름이 있는 요소 보호');
        protectTemplateCheck.value = true;
        
        // 버튼
        var buttonGroup = dialog.add('group');
        var previewBtn = buttonGroup.add('button', undefined, '미리보기');
        var okBtn = buttonGroup.add('button', undefined, '삭제 실행');
        var cancelBtn = buttonGroup.add('button', undefined, '취소');
        
        // 템플릿 요소 보호 목록
        var PROTECTED_NAMES = [
            '페이지', '순서1', '순서2', '순서3', '순서4',
            '상품명1', '상품명2', '상품명3', '상품명4',
            '용량1', '용량2', '용량3', '용량4',
            '원재료1', '원재료2', '원재료3', '원재료4',
            '보관방법1', '보관방법2', '보관방법3', '보관방법4'
        ];
        
        function getTargetLayers() {
            var layers = [];
            
            if (layerRadio1.value) {
                // 활성 레이어만
                if (doc.activeLayer) {
                    layers.push(doc.activeLayer);
                }
            } else if (layerRadio2.value) {
                // auto_layer만
                for (var i = 0; i < doc.layers.length; i++) {
                    if (doc.layers[i].name === 'auto_layer') {
                        layers.push(doc.layers[i]);
                        break;
                    }
                }
            } else {
                // 전체 레이어
                for (var j = 0; j < doc.layers.length; j++) {
                    layers.push(doc.layers[j]);
                }
            }
            
            return layers;
        }
        
        function isProtected(itemName) {
            if (!protectTemplateCheck.value && !protectNamedCheck.value) {
                return false;
            }
            
            if (protectTemplateCheck.value) {
                for (var i = 0; i < PROTECTED_NAMES.length; i++) {
                    if (itemName === PROTECTED_NAMES[i]) {
                        return true;
                    }
                }
            }
            
            if (protectNamedCheck.value && itemName && itemName !== '(이름없음)') {
                return true;
            }
            
            return false;
        }
        
        function findDeletableItems() {
            var layers = getTargetLayers();
            var deletableItems = {
                textFrames: [],
                placedItems: [],
                otherItems: []
            };
            
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];
                
                // 빈 텍스트 프레임
                if (deleteTextCheck.value) {
                    for (var j = 0; j < layer.textFrames.length; j++) {
                        var tf = layer.textFrames[j];
                        var content = (tf.contents || '').replace(/\s/g, '');
                        var itemName = tf.name || '(이름없음)';
                        
                        if (content.length === 0 && !isProtected(itemName)) {
                            deletableItems.textFrames.push({
                                item: tf,
                                name: itemName,
                                layer: layer.name
                            });
                        }
                    }
                }
                
                // 투명도 0인 이미지
                if (deleteImageCheck.value) {
                    for (var k = 0; k < layer.placedItems.length; k++) {
                        var pi = layer.placedItems[k];
                        var itemName = pi.name || '(이름없음)';
                        
                        if (pi.opacity === 0 && !isProtected(itemName)) {
                            deletableItems.placedItems.push({
                                item: pi,
                                name: itemName,
                                layer: layer.name
                            });
                        }
                    }
                }
                
                // 완전 투명한 다른 객체들
                if (deleteInvisibleCheck.value) {
                    for (var l = 0; l < layer.pageItems.length; l++) {
                        var item = layer.pageItems[l];
                        var itemName = item.name || '(이름없음)';
                        
                        if (item.typename !== 'TextFrame' && 
                            item.opacity === 0 && 
                            !isProtected(itemName)) {
                            deletableItems.otherItems.push({
                                item: item,
                                name: itemName,
                                type: item.typename,
                                layer: layer.name
                            });
                        }
                    }
                }
            }
            
            return deletableItems;
        }
        
        previewBtn.onClick = function() {
            var items = findDeletableItems();
            var total = items.textFrames.length + items.placedItems.length + items.otherItems.length;
            
            if (total === 0) {
                alert('삭제할 항목이 없습니다.');
                return;
            }
            
            var preview = '삭제 예정 항목 (' + total + '개):\n\n';
            
            if (items.textFrames.length > 0) {
                preview += '빈 텍스트 프레임 (' + items.textFrames.length + '개):\n';
                for (var i = 0; i < Math.min(items.textFrames.length, 3); i++) {
                    preview += '• ' + items.textFrames[i].name + ' [' + items.textFrames[i].layer + ']\n';
                }
                if (items.textFrames.length > 3) preview += '... 외 ' + (items.textFrames.length - 3) + '개\n';
                preview += '\n';
            }
            
            if (items.placedItems.length > 0) {
                preview += '투명도 0인 이미지 (' + items.placedItems.length + '개):\n';
                for (var j = 0; j < Math.min(items.placedItems.length, 3); j++) {
                    preview += '• ' + items.placedItems[j].name + ' [' + items.placedItems[j].layer + ']\n';
                }
                if (items.placedItems.length > 3) preview += '... 외 ' + (items.placedItems.length - 3) + '개\n';
                preview += '\n';
            }
            
            if (items.otherItems.length > 0) {
                preview += '기타 투명 객체 (' + items.otherItems.length + '개):\n';
                for (var k = 0; k < Math.min(items.otherItems.length, 3); k++) {
                    preview += '• ' + items.otherItems[k].name + ' (' + items.otherItems[k].type + ') [' + items.otherItems[k].layer + ']\n';
                }
                if (items.otherItems.length > 3) preview += '... 외 ' + (items.otherItems.length - 3) + '개\n';
            }
            
            alert(preview);
        };
        
        okBtn.onClick = function() { dialog.close(1); };
        cancelBtn.onClick = function() { dialog.close(0); };
        
        if (dialog.show() === 1) {
            // 삭제 실행
            var items = findDeletableItems();
            var totalDeleted = 0;
            
            // 역순으로 삭제 (인덱스 문제 방지)
            for (var m = items.textFrames.length - 1; m >= 0; m--) {
                items.textFrames[m].item.remove();
                totalDeleted++;
            }
            
            for (var n = items.placedItems.length - 1; n >= 0; n--) {
                items.placedItems[n].item.remove();
                totalDeleted++;
            }
            
            for (var o = items.otherItems.length - 1; o >= 0; o--) {
                items.otherItems[o].item.remove();
                totalDeleted++;
            }
            
            alert('🎉 고급 삭제 완료!\n\n' +
                  '빈 텍스트 프레임: ' + items.textFrames.length + '개\n' +
                  '투명도 0인 이미지: ' + items.placedItems.length + '개\n' +
                  '기타 투명 객체: ' + items.otherItems.length + '개\n\n' +
                  '총 ' + totalDeleted + '개 항목이 삭제되었습니다.\n\n' +
                  '✅ 템플릿 요소는 보호되었습니다.');
        }
        
    } catch (e) {
        alert('오류 발생: ' + e.message);
    }
}

// 실행
advancedDeleteEmptyItems();
