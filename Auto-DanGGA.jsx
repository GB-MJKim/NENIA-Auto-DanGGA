// =====================================
// 네니아 완전 안전 버전 - line.trim() 에러 완전 해결
// CSV 파일 수정 없이 콤마 포함 필드 완벽 처리
// =====================================

// ★ 디버깅 모드 설정
var DEBUG_MODE = false;
var TEMPLATE_LAYER_NAME = "auto_layer";

function log(msg) {
  if (DEBUG_MODE) {
    $.writeln("[DEBUG] " + msg);
  }
}

function debugAlert(msg) {
  if (DEBUG_MODE) {
    alert("[DEBUG] " + msg);
  }
}

// ✅ 안전한 trim 함수 (에러 방지)
function safeTrim(value) {
  if (typeof value !== "string") {
    return "";
  }

  // .trim() 메서드가 있는지 확인
  if (typeof value.trim === "function") {
    return value.trim();
  }

  // 수동으로 앞뒤 공백 제거
  var startIdx = 0;
  var endIdx = value.length - 1;

  while (
    startIdx < value.length &&
    (value.charAt(startIdx) === " " ||
      value.charAt(startIdx) === "\t" ||
      value.charAt(startIdx) === "\n" ||
      value.charAt(startIdx) === "\r")
  ) {
    startIdx++;
  }

  while (
    endIdx >= startIdx &&
    (value.charAt(endIdx) === " " ||
      value.charAt(endIdx) === "\t" ||
      value.charAt(endIdx) === "\n" ||
      value.charAt(endIdx) === "\r")
  ) {
    endIdx--;
  }

  if (startIdx > endIdx) {
    return "";
  }

  var result = "";
  for (var i = startIdx; i <= endIdx; i++) {
    result += value.charAt(i);
  }

  return result;
}

// ✅ 안전한 문자열 길이 체크
function safeLength(value) {
  if (typeof value !== "string") {
    return 0;
  }
  return value.length;
}

// ✅ 메인 함수
function applyNeniaSafeVersion() {
  try {
    debugAlert("1. 네니아 안전 버전 시작");

    var doc = app.activeDocument;
    var templateLayer = getLayerByName(TEMPLATE_LAYER_NAME, doc);

    if (!templateLayer) {
      alert(
        '❌ 템플릿 레이어 "' +
          TEMPLATE_LAYER_NAME +
          '"를 찾을 수 없습니다.\n\n' +
          "해결 방법:\n" +
          "1. 레이어 패널에서 새 레이어 생성\n" +
          '2. 레이어 이름을 "' +
          TEMPLATE_LAYER_NAME +
          '"로 변경\n' +
          "3. 텍스트 프레임 생성 후 이름 설정 (상품명1, 용량1, 원재료1...)\n" +
          "4. 스크립트 다시 실행"
      );
      return;
    }

    debugAlert("2. 템플릿 레이어 발견: " + templateLayer.name);

    var csvFile = File.openDialog("CSV 파일을 선택하세요", "*.csv");
    if (!csvFile) {
      debugAlert("파일 선택 취소");
      return;
    }

    debugAlert("3. 파일 선택: " + csvFile.name);

    csvFile.open("r");
    var content = csvFile.read();
    csvFile.close();

    if (!content || content.length === 0) {
      alert("파일이 비어있습니다.");
      return;
    }

    var lines = safeSplitLinesFixed(content);
    debugAlert("4. 줄 수: " + lines.length);

    var productDataArray = [];

    // ✅ 안전한 CSV 파싱
    for (var i = 2; i < lines.length && productDataArray.length < 4; i++) {
      var line = lines[i];

      // ✅ 안전한 타입 체크와 길이 체크
      if (typeof line === "string" && safeLength(safeTrim(line)) > 0) {
        var cells = parseCSVSafe(line);

        debugAlert("5-" + i + ". 파싱 결과: " + cells.length + "개 셀");
        if (DEBUG_MODE && cells.length >= 4) {
          debugAlert(
            '상품명: "' +
              cells[1] +
              '"\n용량: "' +
              cells[2] +
              '"\n원재료: "' +
              (cells[3].length > 50
                ? cells[3].substring(0, 50) + "..."
                : cells[3]) +
              '"'
          );
        }

        if (cells.length >= 4) {
          productDataArray.push({
            상품명: cells[1] || "",
            용량: cells[2] || "",
            원재료: cells[3] || "",
          });
        }
      }
    }

    if (productDataArray.length === 0) {
      alert("상품 데이터가 없습니다!\n\nCSV 형식을 확인하세요.");
      return;
    }

    debugAlert("6. 추출된 상품 수: " + productDataArray.length);

    // 디버깅 모드에서 텍스트 프레임 테스트
    if (DEBUG_MODE) {
      testTextFrames(templateLayer);
    }

    // 데이터 적용
    for (var p = 0; p < productDataArray.length; p++) {
      var productNumber = p + 1;
      debugAlert(
        "7-" + productNumber + ". 상품 " + productNumber + " 적용 시작"
      );
      updateProductData(productNumber, productDataArray[p], templateLayer);
    }

    alert(
      "🎉 완료!\n" +
        productDataArray.length +
        "개 상품이 성공적으로 적용되었습니다.\n\n✅ line.trim() 에러 완전 해결됨"
    );
  } catch (e) {
    alert("❌ 오류: " + e.message + "\n라인: " + e.line);
    log("상세 오류: " + e.toString());
  }
}

// ✅ 안전한 줄 분할 (line.trim() 사용 안함)
function safeSplitLinesFixed(content) {
  var result = [];
  var current = "";
  var inQuotes = false;

  for (var i = 0; i < content.length; i++) {
    var currentChar = content.charAt(i);

    if (currentChar === '"') {
      inQuotes = !inQuotes;
      current += currentChar;
    } else if ((currentChar === "\n" || currentChar === "\r") && !inQuotes) {
      // ✅ safeTrim 사용 (line.trim() 대신)
      if (safeLength(safeTrim(current)) > 0) {
        result.push(current);
      }
      current = "";

      if (
        currentChar === "\r" &&
        i + 1 < content.length &&
        content.charAt(i + 1) === "\n"
      ) {
        i++;
      }
    } else {
      current += currentChar;
    }
  }

  // ✅ 마지막 줄도 안전하게 처리
  if (safeLength(safeTrim(current)) > 0) {
    result.push(current);
  }

  return result;
}

// ✅ 안전한 CSV 파싱 (콤마 포함 필드 처리)
function parseCSVSafe(line) {
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
      result.push(cleanCellSafe(current));
      current = "";
    } else {
      current += currentChar;
    }

    i++;
  }

  result.push(cleanCellSafe(current));
  return result;
}

// ✅ 안전한 셀 정리
function cleanCellSafe(value) {
  if (typeof value !== "string") {
    return "";
  }

  // ✅ safeTrim 사용
  value = safeTrim(value);

  if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
    value = value.substring(1, value.length - 1);
    value = value.replace(/""/g, '"');
  }

  return value;
}

// ✅ 상품 데이터 업데이트
function updateProductData(productIndex, productData, layer) {
  var fields = [
    {
      name: "상품명" + productIndex,
      value: productData.상품명,
      korean: "상품명",
    },
    { name: "용량" + productIndex, value: productData.용량, korean: "용량" },
    {
      name: "원재료" + productIndex,
      value: productData.원재료,
      korean: "원재료",
    },
  ];

  for (var f = 0; f < fields.length; f++) {
    var field = fields[f];
    var textFrame = findTextInLayer(field.name, layer);

    if (textFrame) {
      var oldContent = textFrame.contents;
      textFrame.contents = field.value;

      log(
        "✅ 성공: " +
          field.name +
          " (" +
          field.korean +
          ') = "' +
          (field.value.length > 30
            ? field.value.substring(0, 30) + "..."
            : field.value) +
          '"'
      );

      if (DEBUG_MODE) {
        alert(
          "✅ 업데이트 성공!\n" +
            "대상: " +
            field.name +
            " (" +
            field.korean +
            ")\n" +
            '이전: "' +
            oldContent +
            '"\n' +
            '변경: "' +
            (field.value.length > 100
              ? field.value.substring(0, 100) + "..."
              : field.value) +
            '"'
        );
      }
    } else {
      log("❌ 실패: " + field.name + " 텍스트 프레임 없음");

      if (DEBUG_MODE) {
        alert(
          "❌ 실패!\n" +
            "대상: " +
            field.name +
            " (" +
            field.korean +
            ")\n" +
            "원인: 텍스트 프레임을 찾을 수 없음\n\n" +
            "해결방법:\n" +
            '1. 레이어 "' +
            layer.name +
            '"에 텍스트 생성\n' +
            '2. 텍스트 프레임 이름을 "' +
            field.name +
            '"로 설정'
        );
      }
    }
  }
}

// ✅ 텍스트 프레임 테스트
function testTextFrames(layer) {
  var testNames = ["상품명1", "용량1", "원재료1", "상품명2"];
  var foundCount = 0;
  var notFound = "";

  for (var i = 0; i < testNames.length; i++) {
    var textFrame = findTextInLayer(testNames[i], layer);
    if (textFrame) {
      foundCount++;
      debugAlert(
        '✅ 찾음: "' + testNames[i] + '" 내용: "' + textFrame.contents + '"'
      );
    } else {
      notFound += testNames[i] + "\n";
    }
  }

  debugAlert(
    "텍스트 프레임 테스트 결과:\n찾음: " +
      foundCount +
      "개\n못 찾음:\n" +
      notFound
  );
}

// ✅ 레이어에서 텍스트 프레임 찾기
function findTextInLayer(name, layer) {
  for (var i = 0; i < layer.pageItems.length; i++) {
    var item = layer.pageItems[i];
    if (item.typename === "TextFrame" && item.name === name) {
      return item;
    }
  }

  // 그룹 내 검색
  function searchGroups(container) {
    for (var j = 0; j < container.groupItems.length; j++) {
      var group = container.groupItems[j];

      for (var k = 0; k < group.textFrames.length; k++) {
        if (group.textFrames[k].name === name) {
          return group.textFrames[k];
        }
      }

      var found = searchGroups(group);
      if (found) return found;
    }
    return null;
  }

  return searchGroups(layer);
}

// ✅ 레이어 찾기
function getLayerByName(layerName, doc) {
  for (var i = 0; i < doc.layers.length; i++) {
    if (doc.layers[i].name === layerName) {
      return doc.layers[i];
    }
  }
  return null;
}

// ✅ 텍스트 프레임 자동 이름 설정 도우미
function setupTextFrameNames() {
  try {
    var doc = app.activeDocument;
    var layer = getLayerByName(TEMPLATE_LAYER_NAME, doc);

    if (!layer) {
      alert('❌ "' + TEMPLATE_LAYER_NAME + '" 레이어 없음');
      return;
    }

    var textFrames = [];
    for (var i = 0; i < layer.pageItems.length; i++) {
      var item = layer.pageItems[i];
      if (item.typename === "TextFrame") {
        textFrames.push(item);
      }
    }

    if (textFrames.length === 0) {
      alert("❌ 텍스트 프레임 없음");
      return;
    }

    var names = [
      "상품명1",
      "용량1",
      "원재료1",
      "상품명2",
      "용량2",
      "원재료2",
      "상품명3",
      "용량3",
      "원재료3",
      "상품명4",
      "용량4",
      "원재료4",
    ];
    var msg = "텍스트 프레임에 이름을 설정하시겠습니까?\n\n";

    for (var j = 0; j < Math.min(textFrames.length, names.length); j++) {
      msg +=
        j + 1 + '. "' + textFrames[j].contents + '" → "' + names[j] + '"\n';
    }

    if (confirm(msg)) {
      for (var k = 0; k < Math.min(textFrames.length, names.length); k++) {
        textFrames[k].name = names[k];
      }
      alert(
        "✅ 완료! " +
          Math.min(textFrames.length, names.length) +
          "개 이름 설정됨"
      );
    }
  } catch (e) {
    alert("오류: " + e.message);
  }
}

// ✅ 스크립트 실행
applyNeniaSafeVersion();

// 텍스트 프레임 이름 설정이 필요하면: setupTextFrameNames();
