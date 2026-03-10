<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <script>
    // 日期格式化
    const weekDays = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
    const today = new Date();
    
    document.getElementById('currentDate').textContent = today.toLocaleDateString('zh-TW', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    });
    document.getElementById('weekday').textContent = weekDays[today.getDay()];

    // 計算總項目數
    const totalItems = document.querySelectorAll('.inspection-item').length;
    document.getElementById('totalCount').textContent = totalItems;

    // 更新進度
    function updateProgress() {
      let checkedCount = 0;
      const items = document.querySelectorAll('.inspection-item');
      
      items.forEach((item, index) => {
        const normalRadio = item.querySelector('input[value="正常"]');
        const abnormalRadio = item.querySelector('input[value="異常"]');
        const noteInput = item.querySelector('.item-note input');
        const noteDiv = item.querySelector('.item-note');
        
        // 更新隱藏的 result 值
        const resultInput = document.getElementById('result_' + index);
        const noteInputHidden = document.getElementById('note_input_' + index);
        
        if (normalRadio.checked) {
          resultInput.value = '正常';
          noteInputHidden.value = '';
          noteDiv.style.display = 'none';
          item.classList.remove('abnormal');
          checkedCount++;
        } else if (abnormalRadio.checked) {
          resultInput.value = '異常';
          noteInputHidden.value = noteInput.value;
          noteDiv.style.display = 'block';
          item.classList.add('abnormal');
          // 異常項目也算完成
          checkedCount++;
        }
      });
      
      document.getElementById('checkedCount').textContent = checkedCount;
      
      const percentage = (checkedCount / totalItems) * 100;
      document.getElementById('progressFill').style.width = percentage + '%';
      
      return checkedCount === totalItems;
    }

    // 提交表單
    function submitForm(event) {
      event.preventDefault();
      
      if (!updateProgress()) {
        alert('請完成所有檢查項目！');
        return;
      }
      
      // 顯示 loading
      document.getElementById('loadingOverlay').classList.add('active');
      document.getElementById('submitBtn').disabled = true;
      
      // 收集表單資料
      const items = [];
      const itemElements = document.querySelectorAll('.inspection-item');
      
      itemElements.forEach((item, index) => {
        const category = item.querySelector('input[name^="category_"]').value;
        const itemName = item.querySelector('input[name^="item_"]').value;
        const result = document.getElementById('result_' + index).value;
        const note = document.getElementById('note_input_' + index).value;
        
        items.push({
          category: category,
          item: itemName,
          result: result,
          note: note
        });
      });
      
      const data = { items: items };
      
      // 发送到 Google Apps Script
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onError)
        .doPost({ postData: { contents: JSON.stringify(data) } });
    }

    // 提交成功
    function onSuccess(result) {
      document.getElementById('loadingOverlay').classList.remove('active');
      document.getElementById('submitBtn').disabled = false;
      
      if (result.success) {
        document.getElementById('successModal').classList.add('active');
      } else {
        document.getElementById('errorMessage').textContent = result.message || '請稍後再試';
        document.getElementById('errorModal').classList.add('active');
      }
    }

    // 提交失敗
    function onError(error) {
      document.getElementById('loadingOverlay').classList.remove('active');
      document.getElementById('submitBtn').disabled = false;
      document.getElementById('errorMessage').textContent = error.message || '請稍後再試';
      document.getElementById('errorModal').classList.add('active');
    }

    // 關閉成功 Modal
    function closeModal() {
      document.getElementById('successModal').classList.remove('active');
      // 重置表單
      location.reload();
    }

    // 關閉錯誤 Modal
    function closeErrorModal() {
      document.getElementById('errorModal').classList.remove('active');
    }

    // 初始化
    document.addEventListener('DOMContentLoaded', function() {
      updateProgress();
    });
  </script>
</head>
<body>
  <!-- JavaScript will be injected by Google Apps Script -->
</body>
</html>
