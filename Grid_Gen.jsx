/*
    POPBOOK / 오토단가용 풀 템플릿 그리드 생성 스크립트

    - 빈 페이지에서 실행
    - auto_layer 레이어 생성
    - Product-1 ~ Product-N 그룹 생성
    - 각 Product 그룹 안에:
      * 전체 박스
      * <Clip Group-img> (실제 클리핑 그룹, 더미 포함 → 9045 에러 방지)
      * BASIC_FIELDS 전부 텍스트 프레임으로 생성
      * NEW / 상품태그 / 인증마크용 그룹들 생성

    단위: pt (1mm ≈ 2.835pt)
*/

(function () {
    if (app.documents.length === 0) {
        alert("열려 있는 문서가 없습니다.");
        return;
    }

    var doc = app.activeDocument;

    // 활성 아트보드 정보
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var ab = doc.artboards[abIndex];
    var abRect = ab.artboardRect; // [left, top, right, bottom]
    var abLeft   = abRect[0];
    var abTop    = abRect[1];
    var abRight  = abRect[2];
    var abBottom = abRect[3];
    var abWidth  = abRight - abLeft;
    var abHeight = abTop   - abBottom;

    // ===== 템플릿에 쓸 필드/태그 이름들 =====
    // 오토단가 CONFIG.BASIC_FIELDS 기준
    var BASIC_FIELDS = [
        "페이지", "순서", "타이틀", "상품명", "용량", "원재료", "보관방법",
        "설명란", "서브단가명", "서브단가", "메인단가명", "메인단가", "개당단가", "알러지"
    ];

    // CONFIG.SPECIAL_ITEMS
    var SPECIAL_ITEMS = ["NEW"];

    // CONFIG.PRODUCT_TAGS
    var PRODUCT_TAGS = ["자연해동", "D-7발주", "개별포장"];

    // CONFIG.CERTIFICATION_MARKS
    var CERT_MARKS = ["HACCP", "유기가공식품", "전통식품", "품질인증",
        "무항생제", "무농약가공식품", "동물복지"];

    // -----------------------------
    // 다이얼로그: 그리드/마진 입력
    // -----------------------------
    var dlg = new Window("dialog", "POPBOOK 상품 그리드 템플릿 만들기");
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    dlg.add("statictext", undefined, "※ 단위: pt (1mm ≈ 2.835pt)");

    var g1 = dlg.add("group");
    g1.add("statictext", undefined, "가로 개수 (columns):");
    var colsEdit = g1.add("edittext", undefined, "4");
    colsEdit.characters = 5;

    var g2 = dlg.add("group");
    g2.add("statictext", undefined, "세로 개수 (rows):");
    var rowsEdit = g2.add("edittext", undefined, "3");
    rowsEdit.characters = 5;

    var g3 = dlg.add("group");
    g3.add("statictext", undefined, "상단 마진 (top margin):");
    var topMarginEdit = g3.add("edittext", undefined, "40");
    topMarginEdit.characters = 7;

    var g4 = dlg.add("group");
    g4.add("statictext", undefined, "상품 간 마진 (between items):");
    var gapEdit = g4.add("edittext", undefined, "20");
    gapEdit.characters = 7;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, "확인", { name: "ok" });
    btnGroup.add("button", undefined, "취소", { name: "cancel" });

    if (dlg.show() !== 1) {
        return; // 취소
    }

    var cols = parseInt(colsEdit.text, 10);
    var rows = parseInt(rowsEdit.text, 10);
    var topMargin = parseFloat(topMarginEdit.text);
    var gap = parseFloat(gapEdit.text);

    if (isNaN(cols) || cols <= 0 ||
        isNaN(rows) || rows <= 0 ||
        isNaN(topMargin) || topMargin < 0 ||
        isNaN(gap) || gap < 0) {
        alert("입력값을 다시 확인해 주세요.\n(모든 값은 0보다 큰 숫자여야 합니다.)");
        return;
    }

    var totalItems = cols * rows;

    // -----------------------------
    // 셀 크기 계산
    // -----------------------------
    var totalGapW = (cols - 1) * gap;
    var totalGapH = (rows - 1) * gap;

    var cellWidth = (abWidth - totalGapW) / cols;
    var cellHeight = (abHeight - topMargin - totalGapH) / rows;

    if (cellWidth <= 0 || cellHeight <= 0) {
        alert("아트보드 크기에 비해 그리드/마진 값이 너무 큽니다.\n값을 줄여서 다시 시도해 주세요.");
        return;
    }

    // -----------------------------
    // auto_layer 레이어 생성/획득
    // -----------------------------
    var autoLayerName = "auto_layer";
    var autoLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === autoLayerName) {
            autoLayer = doc.layers[i];
            break;
        }
    }
    if (!autoLayer) {
        autoLayer = doc.layers.add();
        autoLayer.name = autoLayerName;
    }

    // ===== 유틸: 텍스트 프레임 생성 =====
    function createTextFrame(parentGroup, name, left, top, fontSize) {
        var tf = doc.textFrames.add();
        tf.contents = name;   // 플레이스홀더 텍스트
        tf.name = name;
        tf.left = left;
        tf.top = top;
        try {
            tf.textRange.size = fontSize || 9;
        } catch (e) {}
        tf.move(parentGroup, ElementPlacement.PLACEATEND);
        return tf;
    }

    // ===== 유틸: 동적 태그/인증용 작은 그룹 =====
    function createBadgeGroup(parentGroup, name, left, top, width, height) {
        var g = parentGroup.groupItems.add();
        g.name = name;

        var rect = doc.pathItems.rectangle(top, left, width, height);
        rect.stroked = true;
        rect.strokeWidth = 0.5;
        rect.filled = false;
        rect.move(g, ElementPlacement.PLACEATBEGINNING);

        var tf = doc.textFrames.add();
        tf.contents = name;
        tf.left = left + 2;
        tf.top = top - (height / 2) + 4;
        try { tf.textRange.size = 6; } catch (e) {}
        tf.move(g, ElementPlacement.PLACEATEND);

        g.move(parentGroup, ElementPlacement.PLACEATEND);
        return g;
    }

    // -----------------------------
    // Product 그룹 생성 루프
    // -----------------------------
    var imageAreaRatio = 0.5; // 상단 50%를 이미지 영역으로 사용

    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var index = r * cols + c + 1;

            var productGroup = autoLayer.groupItems.add();
            productGroup.name = "Product-" + index;

            var leftPos = abLeft + c * (cellWidth + gap);
            var topPos  = abTop  - topMargin - r * (cellHeight + gap);

            // 전체 박스
            var boxRect = doc.pathItems.rectangle(topPos, leftPos, cellWidth, cellHeight);
            boxRect.stroked = true;
            boxRect.strokeWidth = 0.5;
            boxRect.filled = false;
            boxRect.move(productGroup, ElementPlacement.PLACEATBEGINNING);

            // ===== 이미지용 클리핑 그룹 <Clip Group-img> =====
            var imgHeight = cellHeight * imageAreaRatio;
            var imgTop = topPos - 14; // 위에서 약간 내려서 여백
            var imgLeft = leftPos + 8;
            var imgWidth = cellWidth - 16;

            var clipGroup = productGroup.groupItems.add();
            clipGroup.name = "<Clip Group-img>";

            var clipRect = doc.pathItems.rectangle(imgTop, imgLeft, imgWidth, imgHeight);
            clipRect.stroked = true;
            clipRect.strokeWidth = 0.5;
            clipRect.filled = false;
            clipRect.clipping = true;
            clipRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);

            // 9045 에러 방지용 더미 객체 (투명 사각형)
            var dummyRect = doc.pathItems.rectangle(imgTop - 5, imgLeft + 5, imgWidth - 10, imgHeight - 10);
            dummyRect.stroked = false;
            dummyRect.filled = true;
            dummyRect.opacity = 0;
            dummyRect.move(clipGroup, ElementPlacement.PLACEATEND);

            clipGroup.clipped = true;
            clipGroup.move(productGroup, ElementPlacement.PLACEATEND);

            // ===== 텍스트 영역 =====
            var textMarginLeft = leftPos + 10;
            var textAreaTop = imgTop - imgHeight - 10;
            var lineGap = 12;

            // 페이지 / 순서 (왼쪽 상단 작게)
            createTextFrame(productGroup, "페이지", leftPos + 6, topPos - 6, 6);
            createTextFrame(productGroup, "순서", leftPos + 40, topPos - 6, 6);

            // 메인 정보: 타이틀 / 상품명 / 용량
            var currentTop = textAreaTop;
            createTextFrame(productGroup, "타이틀", textMarginLeft, currentTop, 9); currentTop -= lineGap;
            createTextFrame(productGroup, "상품명", textMarginLeft, currentTop, 10); currentTop -= lineGap;
            createTextFrame(productGroup, "용량",   textMarginLeft, currentTop, 8);  currentTop -= (lineGap + 4);

            // 설명란 (조금 더 아래, 여러 줄용)
            var descTop = currentTop;
            createTextFrame(productGroup, "설명란", textMarginLeft, descTop, 8);
            currentTop = descTop - (lineGap * 2);

            // 원재료 / 알러지 / 보관방법 (하단 왼쪽)
            var infoLeft = textMarginLeft;
            var infoTop = leftPos + 12; // 나중에 아래에서 다시 세팅
            // 하단 기준으로 계산
            var bottomAreaTop = topPos - cellHeight + 40;

            createTextFrame(productGroup, "원재료", infoLeft, bottomAreaTop, 7);
            createTextFrame(productGroup, "알러지", infoLeft, bottomAreaTop - lineGap, 7);
            createTextFrame(productGroup, "보관방법", infoLeft, bottomAreaTop - lineGap * 2, 7);

            // 가격 영역 (하단 오른쪽 쪽)
            var priceLeft = leftPos + cellWidth - 80;
            var priceTop = bottomAreaTop;

            createTextFrame(productGroup, "서브단가명", priceLeft, priceTop, 7);
            createTextFrame(productGroup, "서브단가",   priceLeft, priceTop - lineGap, 9);
            createTextFrame(productGroup, "메인단가명", priceLeft, priceTop - lineGap * 2, 7);
            createTextFrame(productGroup, "메인단가",   priceLeft, priceTop - lineGap * 3, 11);
            createTextFrame(productGroup, "개당단가",   priceLeft, priceTop - lineGap * 4, 7);

            // ===== 동적 요소: NEW / 태그 / 인증마크 =====
            var badgeWidth = 36;
            var badgeHeight = 10;
            var badgeGap = 2;

            // NEW (맨 위 오른쪽)
            var badgeLeft = leftPos + cellWidth - badgeWidth - 8;
            var badgeTop = topPos - 6;
            for (var si = 0; si < SPECIAL_ITEMS.length; si++) {
                createBadgeGroup(productGroup, SPECIAL_ITEMS[si],
                    badgeLeft, badgeTop, badgeWidth, badgeHeight);
                badgeTop -= (badgeHeight + badgeGap);
            }

            // 상품 태그 (자연해동 / D-7발주 / 개별포장)
            badgeTop = imgTop + 4;
            for (var ti = 0; ti < PRODUCT_TAGS.length; ti++) {
                createBadgeGroup(productGroup, PRODUCT_TAGS[ti],
                    badgeLeft, badgeTop, badgeWidth, badgeHeight);
                badgeTop -= (badgeHeight + badgeGap);
            }

            // 인증마크 (HACCP 등) — 이미지 대신 작은 박스로 자리만 잡아줌
            var certLeft = leftPos + cellWidth - badgeWidth - 8;
            var certTop = bottomAreaTop - 4;
            for (var ci = 0; ci < CERT_MARKS.length; ci++) {
                createBadgeGroup(productGroup, CERT_MARKS[ci],
                    certLeft, certTop, badgeWidth, badgeHeight);
                certTop -= (badgeHeight + badgeGap);
            }

            // 그룹 위치 정렬
            productGroup.position = [leftPos, topPos];
        }
    }

    alert(
        "auto_layer에 Product-1 ~ Product-" +
        totalItems + " 템플릿을 생성했습니다.\n\n" +
        "BASIC_FIELDS 전체 + NEW/태그/인증마크 그룹이 모두 포함되어 있습니다."
    );
})();
