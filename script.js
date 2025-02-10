document.addEventListener("DOMContentLoaded", function () {
  const numStudentsInput = document.getElementById("numStudents");
  const rowsInput = document.getElementById("rows");
  const colsInput = document.getElementById("cols");
  const generateButton = document.getElementById("generateSeats");
  const randomizeButton = document.getElementById("randomize");
  const seatingChart = document.getElementById("seatingChart");
  const printBtn = document.getElementById("printBtn");
  const fileInput = document.getElementById("fileInput");
  const importBtn = document.getElementById("importBtn");
  const fileName = document.getElementById("fileName");
  const columnSelectModal = document.getElementById("columnSelectModal");
  const columnList = document.getElementById("columnList");
  const confirmColumnBtn = document.getElementById("confirmColumn");
  const cancelColumnBtn = document.getElementById("cancelColumn");

  let seats = [];
  let dragSrcEl = null;
  let importedNames = [];
  let excelData = null;
  let selectedCells = new Set(); // 存储选中的单元格
  let isMouseDown = false; // 跟踪鼠标按下状态
  let isSelecting = null; // 跟踪是选中还是取消选中的操作
  let lastCell = null;
  let startCell = null; // 记录选择起始单元格
  let selectedColumns = [];

  // 添加右键菜单相关代码
  let contextMenu;

  // 加载保存的座位信息
  function loadSavedSeats() {
    const savedData = localStorage.getItem("seatingData");
    if (savedData) {
      const data = JSON.parse(savedData);
      numStudentsInput.value = data.numStudents;
      rowsInput.value = data.rows;
      colsInput.value = data.cols;

      // 只加载保存的名字，不补空
      importedNames = data.names;

      // 生成座位
      generateSeats();
    }
  }

  // 保存座位信息
  function saveSeatingData() {
    const seatingData = {
      numStudents: numStudentsInput.value,
      rows: rowsInput.value,
      cols: colsInput.value,
      // 只保存有名字的座位
      names: Array.from(document.querySelectorAll(".seat:not(.empty) input"))
        .map((input) => input.value)
        .filter((name) => name.trim() !== ""),
    };
    localStorage.setItem("seatingData", JSON.stringify(seatingData));
  }

  // 在页面加载时恢复座位
  loadSavedSeats();

  // 绑定生成座位按钮的点击事件
  generateButton.addEventListener("click", generateSeats);

  function handleDragStart(e) {
    this.style.opacity = "0.4";
    dragSrcEl = this;
    e.dataTransfer.effectAllowed = "move";
  }

  function handleDragEnd(e) {
    this.style.opacity = "1";
    document.querySelectorAll(".seat").forEach((seat) => {
      seat.classList.remove("drag-over");
    });
  }

  function handleDragOver(e) {
    e.preventDefault();
    return false;
  }

  function handleDragEnter(e) {
    this.classList.add("drag-over");
  }

  function handleDragLeave(e) {
    this.classList.remove("drag-over");
  }

  function handleDrop(e) {
    e.stopPropagation();

    if (dragSrcEl !== this) {
      const srcInput = dragSrcEl.querySelector("input");
      const destInput = this.querySelector("input");

      // 如果目标是空座位
      if (this.classList.contains("empty")) {
        // 移动到空座位
        destInput.value = srcInput.value;
        destInput.disabled = false;
        this.dataset.value = srcInput.value;
        srcInput.value = "";
        dragSrcEl.dataset.value = "";
        srcInput.disabled = true;

        // 更新座位状态
        this.classList.remove("empty");
        dragSrcEl.classList.add("empty");

        // 更新占位符
        destInput.placeholder = `座位 ${
          Array.from(seatingChart.children).indexOf(this) + 1
        }`;
        srcInput.placeholder = "空座";
      } else {
        // 普通座位之间交换
        const tempValue = srcInput.value;
        srcInput.value = destInput.value;
        destInput.value = tempValue;

        const tempDataValue = dragSrcEl.dataset.value;
        dragSrcEl.dataset.value = this.dataset.value;
        this.dataset.value = tempDataValue;
      }

      // 每次调整座位后保存
      saveSeatingData();
    }

    return false;
  }

  function removeDragListeners(seat) {
    seat.removeAttribute("draggable");
    seat.removeEventListener("dragstart", handleDragStart);
    seat.removeEventListener("dragend", handleDragEnd);
    seat.removeEventListener("dragover", handleDragOver);
    seat.removeEventListener("dragenter", handleDragEnter);
    seat.removeEventListener("dragleave", handleDragLeave);
    seat.removeEventListener("drop", handleDrop);
  }

  function addDragListeners(seat) {
    seat.setAttribute("draggable", true);
    seat.addEventListener("dragstart", handleDragStart);
    seat.addEventListener("dragend", handleDragEnd);
    seat.addEventListener("dragover", handleDragOver);
    seat.addEventListener("dragenter", handleDragEnter);
    seat.addEventListener("dragleave", handleDragLeave);
    seat.addEventListener("drop", handleDrop);

    // 添加右键菜单事件
    seat.addEventListener("contextmenu", function (e) {
      e.preventDefault();
      e.stopPropagation();

      if (!contextMenu) {
        contextMenu = createContextMenu();
      }

      // 根据座位状态显示/隐藏菜单项
      const deleteItem = contextMenu.querySelector('[data-action="delete"]');
      const clearItem = contextMenu.querySelector('[data-action="clear"]');

      if (this.classList.contains("empty")) {
        deleteItem.style.display = "flex";
        clearItem.style.display = "none";
      } else {
        deleteItem.style.display = "flex";
        clearItem.style.display = "flex";
      }

      // 获取被点击座位相对于视口的位置
      const seatRect = this.getBoundingClientRect();
      const menuWidth = 180;
      const menuHeight = 120;
      const gap = 5;

      // 计算菜单位置（相对于视口）
      let x = seatRect.left + (seatRect.width - menuWidth) / 2;
      let y = seatRect.bottom + gap;

      // 确保菜单不会超出视口边界
      x = Math.max(5, Math.min(x, window.innerWidth - menuWidth - 5));
      y = Math.min(y, window.innerHeight - menuHeight - 5);

      // 设置菜单位置（使用视口坐标）
      contextMenu.style.left = `${x}px`;
      contextMenu.style.top = `${y}px`;
      contextMenu.classList.remove("hidden");

      // 记录当前右键的座位
      contextMenu.dataset.targetSeat = Array.from(
        seatingChart.children
      ).indexOf(this);
    });
  }

  // 导入按钮点击事件
  importBtn.addEventListener("click", function () {
    // 创建一个新的文件输入元素
    const newFileInput = document.createElement("input");
    newFileInput.type = "file";
    newFileInput.accept = ".xlsx,.xls,.csv";
    newFileInput.style.display = "none";

    // 复制原始文件输入的事件处理程序
    newFileInput.addEventListener("change", function (e) {
      const file = e.target.files[0];
      if (!file) return;

      fileName.textContent = file.name;
      const reader = new FileReader();

      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

          // 添加调试信息，确保数据被正确解析
          console.log("Excel data:", excelData);

          if (excelData && excelData.length > 0) {
            const columnCount = excelData[0].length;
            showColumnSelectModal(new Array(columnCount).fill(null));
            selectedCells.clear();
          } else {
            alert("文件中没有数据");
          }
        } catch (error) {
          console.error("Error reading Excel:", error);
          alert("文件读取失败，请确保文件格式正确");
        }
      };

      reader.readAsArrayBuffer(file);

      // 移除临时创建的输入元素
      document.body.removeChild(newFileInput);
    });

    // 添加到文档并触发点击
    document.body.appendChild(newFileInput);
    newFileInput.click();
  });

  // 显示列选择弹窗
  function showColumnSelectModal(columns) {
    columnList.innerHTML = "";

    // 创建表格视图
    const table = document.createElement("table");
    table.className = "excel-table";

    // 创建表格内容
    const tbody = document.createElement("tbody");
    excelData.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      row.forEach((cell, colIndex) => {
        const td = document.createElement("td");
        if (cell) {
          td.textContent = cell;
          td.dataset.value = cell;
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    columnList.appendChild(table);

    // 添加调整大小的功能
    const modalContent = columnSelectModal.querySelector(".modal-content");
    const resizeHandle = document.createElement("div");
    resizeHandle.className = "resize-handle";
    modalContent.appendChild(resizeHandle);

    let isResizing = false;
    let startX, startY, startWidth, startHeight;

    resizeHandle.addEventListener("mousedown", (e) => {
      isResizing = true;
      startX = e.clientX;
      startY = e.clientY;
      startWidth = modalContent.offsetWidth;
      startHeight = modalContent.offsetHeight;

      document.addEventListener("mousemove", handleMouseMove);
      document.addEventListener("mouseup", handleMouseUp);
    });

    function handleMouseMove(e) {
      if (!isResizing) return;

      const newWidth = startWidth + (e.clientX - startX);
      const newHeight = startHeight + (e.clientY - startY);

      // 限制最小尺寸
      modalContent.style.width = Math.max(300, newWidth) + "px";
      modalContent.style.height = Math.max(400, newHeight) + "px";
    }

    function handleMouseUp() {
      isResizing = false;
      document.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
    }

    // 选择功能
    let isSelecting = false;
    let isAddingSelection = null;
    let lastSelectedCell = null;

    // 鼠标事件处理
    table.addEventListener("mousedown", (e) => {
      if (e.target.tagName === "TD") {
        isSelecting = true;
        lastSelectedCell = e.target;
        isAddingSelection = !e.target.classList.contains("selected");

        if (isAddingSelection) {
          e.target.classList.add("selected");
        } else {
          e.target.classList.remove("selected");
        }
      }
    });

    table.addEventListener("mousemove", (e) => {
      if (!isSelecting || e.target.tagName !== "TD") return;

      if (e.target !== lastSelectedCell) {
        if (isAddingSelection) {
          e.target.classList.add("selected");
        } else {
          e.target.classList.remove("selected");
        }
        lastSelectedCell = e.target;
      }
    });

    document.addEventListener("mouseup", () => {
      isSelecting = false;
      lastSelectedCell = null;
    });

    // 修改处理重复名字的逻辑
    function showDuplicateNamesModal(duplicateNames, callback) {
      const modalHTML = `
        <div id="duplicateNamesModal" class="modal">
          <div class="modal-content">
            <div class="modal-header">
              <h2 class="modal-title">发现重复的名字</h2>
              <button class="close-button">&times;</button>
            </div>
            <div class="modal-body">
              <p>以下名字已存在，是否继续添加？</p>
              <div class="duplicate-list">
                ${duplicateNames
                  .map(
                    (name) => `
                  <div class="duplicate-item">
                    <label>
                      <input type="checkbox" value="${name}" checked>
                      ${name}
                    </label>
                  </div>
                `
                  )
                  .join("")}
              </div>
            </div>
            <div class="modal-footer">
              <button id="keepSelected" class="confirm-btn">保留选中</button>
              <button id="skipAll" class="secondary-btn">全部跳过</button>
            </div>
          </div>
        </div>
      `;

      // 如果已存在重复名字弹窗，先移除
      const existingModal = document.getElementById("duplicateNamesModal");
      if (existingModal) {
        existingModal.remove();
      }

      // 添加弹窗到页面
      document.body.insertAdjacentHTML("beforeend", modalHTML);

      const modal = document.getElementById("duplicateNamesModal");
      const keepSelectedBtn = document.getElementById("keepSelected");
      const skipAllBtn = document.getElementById("skipAll");
      const closeBtn = modal.querySelector(".close-button");

      // 移除之前的事件监听器（如果有的话）
      const newKeepSelectedBtn = keepSelectedBtn.cloneNode(true);
      const newSkipAllBtn = skipAllBtn.cloneNode(true);
      const newCloseBtn = closeBtn.cloneNode(true);

      keepSelectedBtn.parentNode.replaceChild(
        newKeepSelectedBtn,
        keepSelectedBtn
      );
      skipAllBtn.parentNode.replaceChild(newSkipAllBtn, skipAllBtn);
      closeBtn.parentNode.replaceChild(newCloseBtn, closeBtn);

      // 重新添加事件监听器
      newKeepSelectedBtn.addEventListener("click", () => {
        const selectedNames = Array.from(
          modal.querySelectorAll("input:checked")
        ).map((input) => input.value);
        modal.remove();
        callback(selectedNames);
        document.querySelectorAll(".modal").forEach((modal) => {
          modal.classList.add("hidden");
        });
      });

      newSkipAllBtn.addEventListener("click", () => {
        modal.remove();
        callback([]);
        document.querySelectorAll(".modal").forEach((modal) => {
          modal.classList.add("hidden");
        });
      });

      newCloseBtn.addEventListener("click", () => {
        modal.remove();
        callback([]);
        document.querySelectorAll(".modal").forEach((modal) => {
          modal.classList.add("hidden");
        });
      });

      // 显示弹窗
      modal.classList.remove("hidden");
    }

    // 修改确认按钮的处理逻辑
    const confirmBtn = document.getElementById("confirmColumn");
    confirmBtn.addEventListener("click", () => {
      const selectedCells = tbody.querySelectorAll("td.selected");
      const selectedNames = Array.from(selectedCells)
        .map((cell) => cell.textContent.trim())
        .filter((name) => name !== "");

      if (selectedNames.length > 0) {
        // 获取当前已有的名字
        const existingNames = Array.from(
          document.querySelectorAll(".seat:not(.empty) input")
        )
          .map((input) => input.value.trim())
          .filter((name) => name !== "");

        // 检查重复的名字
        const duplicates = selectedNames.filter((name) =>
          existingNames.includes(name)
        );

        if (duplicates.length > 0) {
          // 显示重复名字提示
          showDuplicateNamesModal(duplicates, (namesToKeep) => {
            // 过滤掉不需要保留的名字
            const finalNames = selectedNames.filter(
              (name) => !duplicates.includes(name) || namesToKeep.includes(name)
            );

            importedNames = [...existingNames, ...finalNames];
            processImportedNames();
          });
        } else {
          importedNames = [...existingNames, ...selectedNames];
          processImportedNames();
          columnSelectModal.classList.add("hidden");
        }
      }
    });

    // 添加关闭按钮事件处理
    const closeBtn = document.getElementById("closeColumnModal");
    const oldCloseBtn = closeBtn.cloneNode(true);
    closeBtn.parentNode.replaceChild(oldCloseBtn, closeBtn);

    oldCloseBtn.addEventListener("click", () => {
      columnSelectModal.classList.add("hidden");
    });

    // 添加点击遮罩层关闭功能
    columnSelectModal.addEventListener("click", (e) => {
      if (e.target === columnSelectModal) {
        columnSelectModal.classList.add("hidden");
      }
    });

    // 显示弹窗
    columnSelectModal.classList.remove("hidden");
  }

  // 取消按钮点击事件
  cancelColumnBtn.addEventListener("click", hideColumnSelectModal);

  // 当输入人数时自动计算推荐的行列数
  numStudentsInput.addEventListener("input", function () {
    const numStudents = parseInt(this.value) || 0;
    if (numStudents > 0) {
      const sqrt = Math.sqrt(numStudents);
      const recommendedCols = Math.ceil(sqrt);
      const recommendedRows = Math.ceil(numStudents / recommendedCols);

      rowsInput.value = recommendedRows;
      colsInput.value = recommendedCols;
    }
  });

  // 修改生成座位的函数
  function generateSeats() {
    const rows = parseInt(rowsInput.value) || 0;
    const cols = parseInt(colsInput.value) || 0;

    if (rows < 1 || cols < 1) {
      alert("请输入有效的行列数！");
      return;
    }

    // 获取要使用的名字列表（优先使用导入的名字，否则使用现有座位的名字）
    let namesToUse = [];
    if (importedNames && importedNames.length > 0) {
      // 使用导入的名字
      namesToUse = importedNames;
    } else {
      // 使用现有座位的名字
      const existingSeats = Array.from(
        document.querySelectorAll(".seat:not(.empty) input")
      );
      namesToUse = existingSeats
        .map((input) => input.value)
        .filter((name) => name.trim() !== "");
    }

    // 检查新的座位数是否足够
    if (namesToUse.length > rows * cols) {
      const result = confirm(
        `当前有 ${namesToUse.length} 个名字，但新的座位表只有 ${
          rows * cols
        } 个位置。\n` +
          `继续操作将会删除多余的名字。\n\n` +
          `是否继续？`
      );

      if (!result) {
        // 用户取消操作，恢复之前的行列数
        const savedData = JSON.parse(localStorage.getItem("seatingData"));
        if (savedData) {
          rowsInput.value = savedData.rows;
          colsInput.value = savedData.cols;
        }
        return;
      }
      // 如果用户确认，截取需要的名字数量
      namesToUse = namesToUse.slice(0, rows * cols);
    }

    // 显示整个教室布局
    document.getElementById("classroom-layout").classList.remove("hidden");

    // 设置网格布局
    seatingChart.style.gridTemplateColumns = `repeat(${cols}, 100px)`;

    // 清空现有座位
    seatingChart.innerHTML = "";
    seats = [];

    // 生成座位
    for (let i = 0; i < rows * cols; i++) {
      const seat = document.createElement("div");
      const input = document.createElement("input");
      input.type = "text";
      input.placeholder = `座位 ${i + 1}`;

      // 如果这个位置有名字，则填入名字
      if (i < namesToUse.length) {
        seat.className = "seat";
        input.value = namesToUse[i];
        seat.dataset.value = namesToUse[i];
      } else {
        seat.className = "seat empty";
        input.disabled = false;
        input.placeholder = `座位 ${i + 1}`;
      }

      // 添加输入事件监听器
      input.addEventListener("input", function () {
        if (this.value.trim() !== "") {
          seat.classList.remove("empty");
          seat.classList.add("editing"); // 添加编辑状态的样式
        } else {
          seat.classList.add("empty");
          seat.classList.remove("editing");
        }
        seat.dataset.value = this.value;
        saveSeatingData();
      });

      // 添加焦点事件监听器
      input.addEventListener("focus", function () {
        seat.classList.add("editing");
      });

      input.addEventListener("blur", function () {
        seat.classList.remove("editing");
        if (!this.value.trim()) {
          seat.classList.add("empty");
        }
      });

      seat.appendChild(input);
      seatingChart.appendChild(seat);

      // 为所有座位添加拖拽功能
      addDragListeners(seat);
      if (!seat.classList.contains("empty")) {
        seats.push(seat);
      }
    }

    // 显示随机调整按钮
    randomizeButton.classList.remove("hidden");

    // 保存数据
    saveSeatingData();

    // 确保在生成座位后初始化右键菜单
    if (!contextMenu) {
      contextMenu = createContextMenu();
    }
  }

  // 在输入框值改变时保存数据
  document.querySelectorAll(".seat input").forEach((input) => {
    input.addEventListener("change", saveSeatingData);
  });

  // 修改随机调整座位的函数
  function randomizeSeats() {
    // 获取所有座位，包括空座位
    const allSeats = Array.from(document.querySelectorAll(".seat"));

    // 获取所有有名字的座位的值
    const names = Array.from(
      document.querySelectorAll(".seat:not(.empty) input")
    )
      .map((input) => input.value)
      .filter((name) => name.trim() !== "");

    // 打乱名字数组
    const shuffledNames = [...names].sort(() => Math.random() - 0.5);

    // 随机选择位置分配名字
    const selectedPositions = new Set();
    let nameIndex = 0;

    while (nameIndex < shuffledNames.length) {
      // 随机选择一个位置
      let position;
      do {
        position = Math.floor(Math.random() * allSeats.length);
      } while (selectedPositions.has(position));

      selectedPositions.add(position);

      const seat = allSeats[position];
      const input = seat.querySelector("input");
      const name = shuffledNames[nameIndex];

      // 更新座位状态
      if (seat.classList.contains("empty")) {
        seat.classList.remove("empty");
        input.disabled = false;
      }

      input.value = name;
      seat.dataset.value = name;
      input.placeholder = `座位 ${position + 1}`;

      nameIndex++;
    }

    // 将剩余的座位设为空座位
    allSeats.forEach((seat, index) => {
      if (!selectedPositions.has(index)) {
        const input = seat.querySelector("input");
        seat.classList.add("empty");
        input.disabled = true;
        input.value = "";
        seat.dataset.value = "";
        input.placeholder = "空座";
      }
    });

    // 保存更新后的座位信息
    saveSeatingData();
  }

  // 绑定随机按钮点击事件
  randomizeButton.addEventListener("click", randomizeSeats);

  // 修改打印按钮的处理函数
  printBtn.addEventListener("click", function () {
    const layout = document.getElementById("classroom-layout");
    if (layout.classList.contains("hidden")) {
      alert("请先生成座位表");
      return;
    }

    // 检查是否有空座位
    const emptySeats = document.querySelectorAll(".seat.empty");
    if (emptySeats.length > 0) {
      // 显示打印选项对话框
      document.getElementById("printOptionModal").classList.remove("hidden");
    } else {
      // 没有空座位，直接打印
      window.print();
    }
  });

  // 确认打印按钮事件
  document
    .getElementById("confirmPrint")
    .addEventListener("click", function () {
      const printEmptySeats =
        document.getElementById("printEmptySeats").checked;
      const emptySeats = document.querySelectorAll(".seat.empty");

      // 添加打印时的类名，用于控制空座位显示
      document.body.classList.add("printing");
      if (!printEmptySeats) {
        document.body.classList.add("hide-empty-seats");
      }

      // 隐藏打印选项对话框
      document.getElementById("printOptionModal").classList.add("hidden");

      // 打印
      window.print();

      // 移除打印时添加的类名
      setTimeout(() => {
        document.body.classList.remove("printing", "hide-empty-seats");
      }, 100);
    });

  // 取消打印按钮事件
  document.getElementById("cancelPrint").addEventListener("click", function () {
    document.getElementById("printOptionModal").classList.add("hidden");
  });

  // 添加清除数据的功能（可选）
  window.clearSeatingData = function () {
    localStorage.removeItem("seatingData");
    location.reload();
  };

  // 重新开始功能
  const restartBtn = document.getElementById("restartBtn");
  restartBtn.addEventListener("click", function () {
    if (confirm("确定要重新开始吗？这将清空所有座位信息。")) {
      // 清空所有输入
      numStudentsInput.value = "";
      rowsInput.value = "";
      colsInput.value = "";
      importedNames = [];
      selectedColumns = []; // 清除选择的列

      // 清空所有状态
      const seats = document.querySelectorAll(".seat");
      seats.forEach((seat) => {
        const input = seat.querySelector("input");
        if (input) {
          input.value = "";
          seat.dataset.value = "";
        }
      });

      // 隐藏座位表和随机按钮
      document.getElementById("classroom-layout").classList.add("hidden");
      randomizeButton.classList.add("hidden");

      // 清空本地存储
      localStorage.removeItem("seatingData");
      localStorage.removeItem("lastSelectedColumns"); // 清除上次选择的记录

      // 清空文件选择
      fileInput.value = "";
      fileName.textContent = "";
    }
  });

  // 将原有的处理逻辑移到单独的函数中
  function processImportedNames() {
    // 如果已经生成了座位表，直接更新现有座位
    const layout = document.getElementById("classroom-layout");
    if (!layout.classList.contains("hidden")) {
      const allInputs = document.querySelectorAll(".seat:not(.empty) input");
      const currentSeats = allInputs.length;

      // 如果导入的名字数量超过当前座位数，自动增加座位
      if (importedNames.length > currentSeats) {
        // 保存当前的行列数
        const currentRows = parseInt(rowsInput.value);
        const currentCols = parseInt(colsInput.value);

        // 计算需要的新行数
        const neededSeats = importedNames.length;
        const newRows = Math.ceil(neededSeats / currentCols);

        // 更新行数和人数
        rowsInput.value = newRows;
        numStudentsInput.value = neededSeats;

        // 重新生成座位表
        generateSeats();

        alert(
          `座位数不足，已自动增加行数到 ${newRows} 行，以容纳所有 ${neededSeats} 个名字`
        );
      } else {
        // 只更新现有座位的内容，不改变座位数量
        allInputs.forEach((input, index) => {
          if (index < importedNames.length) {
            input.value = importedNames[index];
            input.parentElement.dataset.value = importedNames[index];
          }
        });

        alert(`成功导入 ${importedNames.length} 个名字`);
      }
    } else {
      // 自动设置人数和推荐的行列数
      const numStudents = importedNames.length;
      numStudentsInput.value = numStudents;

      // 计算推荐的行列数
      const sqrt = Math.sqrt(numStudents);
      const recommendedCols = Math.ceil(sqrt);
      const recommendedRows = Math.ceil(numStudents / recommendedCols);

      rowsInput.value = recommendedRows;
      colsInput.value = recommendedCols;

      // 生成座位表
      generateSeats();

      alert(`成功导入 ${importedNames.length} 个名字并生成座位表`);
    }
  }

  // 临时选择矩形区域的函数
  function selectRectangleTemp(table, start, end, isSelecting) {
    const startRow = Math.min(
      start.parentElement.rowIndex,
      end.parentElement.rowIndex
    );
    const endRow = Math.max(
      start.parentElement.rowIndex,
      end.parentElement.rowIndex
    );
    const startCol = Math.min(start.cellIndex, end.cellIndex);
    const endCol = Math.max(start.cellIndex, end.cellIndex);

    const rows = table.getElementsByTagName("tr");
    for (let i = startRow; i <= endRow; i++) {
      for (let j = startCol; j <= endCol; j++) {
        const cell = rows[i].cells[j];
        if (cell && cell.dataset.value) {
          cell.classList.add("temp-selected");
        }
      }
    }
  }

  // 隐藏列选择弹窗
  function hideColumnSelectModal() {
    columnSelectModal.classList.add("hidden");
    selectedCells.clear(); // 清空选中状态
  }

  // 修改 createColumnSelectionUI 函数
  function createColumnSelectionUI(headers) {
    const columnList = document.getElementById("columnList");
    columnList.innerHTML = "";

    // 创建表格来显示Excel的列
    const table = document.createElement("table");
    table.classList.add("column-selection-table");

    // 添加表头行
    const headerRow = document.createElement("tr");
    headers.forEach((header, index) => {
      const th = document.createElement("th");
      th.textContent = String.fromCharCode(65 + index); // A, B, C...
      headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // 添加内容行
    const contentRow = document.createElement("tr");
    headers.forEach((header, index) => {
      const td = document.createElement("td");
      td.textContent = header;
      td.dataset.index = index;
      // 如果这个列之前被选中过，添加选中状态
      if (selectedColumns.includes(index)) {
        td.classList.add("selected");
      }
      contentRow.appendChild(td);
    });
    table.appendChild(contentRow);

    columnList.appendChild(table);

    // 使用事件委托处理选择
    let isSelecting = false;
    let startCell = null;
    let isAddingSelection = null;
    let lastSelectedCell = null;
    let lastProcessedIndex = null;
    let selectedPath = new Set(); // 记录鼠标经过的路径

    table.addEventListener("mousedown", (e) => {
      if (e.target.tagName === "TD") {
        isSelecting = true;
        startCell = e.target;
        lastProcessedIndex = parseInt(e.target.dataset.index);
        isAddingSelection = !e.target.classList.contains("selected");
        selectedPath.clear(); // 清空路径记录

        // 记录并处理起始点
        const startIndex = parseInt(e.target.dataset.index);
        if (!isNaN(startIndex)) {
          selectedPath.add(startIndex);
          if (isAddingSelection) {
            e.target.classList.add("selected");
          } else {
            e.target.classList.remove("selected");
          }
        }
      }
    });

    table.addEventListener("mousemove", (e) => {
      if (!isSelecting || e.target.tagName !== "TD") return;

      const currentIndex = parseInt(e.target.dataset.index);
      if (isNaN(currentIndex) || currentIndex === lastProcessedIndex) return;

      // 计算从上一个位置到当前位置的路径
      const direction = currentIndex > lastProcessedIndex ? 1 : -1;
      let index = lastProcessedIndex;

      // 处理鼠标移动路径上的所有单元格
      do {
        index += direction;
        const cell = contentRow.querySelector(`td[data-index="${index}"]`);
        if (cell) {
          selectedPath.add(index);
          if (isAddingSelection) {
            cell.classList.add("selected");
          } else {
            cell.classList.remove("selected");
          }
        }
      } while (index !== currentIndex);

      lastProcessedIndex = currentIndex;
    });

    const mouseUpHandler = () => {
      if (isSelecting) {
        const selectedCells = contentRow.querySelectorAll("td.selected");
        selectedColumns = Array.from(selectedCells)
          .map((cell) => parseInt(cell.dataset.index))
          .sort((a, b) => a - b);

        localStorage.setItem(
          "lastSelectedColumns",
          JSON.stringify(selectedColumns)
        );
      }
      isSelecting = false;
      startCell = null;
      isAddingSelection = null;
      lastSelectedCell = null;
      lastProcessedIndex = null;
      selectedPath.clear();
    };

    document.removeEventListener("mouseup", mouseUpHandler);
    document.addEventListener("mouseup", mouseUpHandler);

    // 防止拖拽过程中选中文本
    table.addEventListener("selectstart", (e) => {
      if (isSelecting) {
        e.preventDefault();
      }
    });
  }

  // 在文件导入时，恢复上次的选择
  function handleFileSelect(event) {
    // 移除选择列相关的代码和提示
    const lastSelection = localStorage.getItem("lastSelectedColumns");
    if (lastSelection) {
      selectedColumns = JSON.parse(lastSelection);
    } else {
      selectedColumns = [];
    }

    // 直接创建选择界面
    createColumnSelectionUI(headers);
  }

  // 修改弹窗的 HTML 结构
  function createModalHTML() {
    return `
      <div id="columnSelectModal" class="modal hidden">
        <div class="modal-content">
          <div class="modal-header">
            <h2 class="modal-title">选择名单</h2>
            <button id="closeColumnModal" class="close-button">&times;</button>
          </div>
          <div id="columnList" class="column-list"></div>
          <div class="modal-footer">
            <button id="confirmColumn" class="confirm-btn">确认</button>
          </div>
        </div>
      </div>
    `;
  }

  // 创建右键菜单
  function createContextMenu() {
    // 如果已存在菜单，先移除
    const existingMenu = document.querySelector(".context-menu");
    if (existingMenu) {
      existingMenu.remove();
    }

    const menu = document.createElement("div");
    menu.className = "context-menu hidden";
    menu.innerHTML = `
      <div class="menu-title">座位操作</div>
      <div class="menu-item" data-action="delete">
        <span class="material-icons">delete</span>
        删除座位
      </div>
      <div class="menu-item" data-action="clear">
        <span class="material-icons">person_remove</span>
        清除名字
      </div>
    `;

    document.body.appendChild(menu);

    // 绑定菜单点击事件
    menu.addEventListener("click", function (e) {
      const menuItem = e.target.closest(".menu-item");
      if (!menuItem) return;

      const action = menuItem.dataset.action;
      const targetSeatIndex = parseInt(menu.dataset.targetSeat);
      const targetSeat = seatingChart.children[targetSeatIndex];

      if (targetSeat) {
        const input = targetSeat.querySelector("input");

        if (action === "delete") {
          if (targetSeat.classList.contains("empty")) {
            // 如果是空座位，直接移除
            targetSeat.classList.add("deleting");
            setTimeout(() => {
              // 直接移除该座位
              targetSeat.remove();

              // 保存更新后的数据
              saveSeatingData();
            }, 300);
          } else {
            // 如果是有名字的座位，设为空座位
            targetSeat.classList.add("deleting");
            setTimeout(() => {
              targetSeat.classList.remove("deleting");
              targetSeat.classList.add("empty");
              input.disabled = true;
              input.value = "";
              targetSeat.dataset.value = "";
              input.placeholder = "空座";
              saveSeatingData();
            }, 300);
          }
        } else if (action === "clear") {
          targetSeat.classList.add("clearing");
          setTimeout(() => {
            targetSeat.classList.remove("clearing");
            targetSeat.classList.add("empty");
            input.disabled = false;
            input.value = "";
            targetSeat.dataset.value = "";
            input.placeholder = `座位 ${targetSeatIndex + 1}`;
            saveSeatingData();
          }, 300);
        }
      }

      menu.classList.add("hidden");
    });

    return menu;
  }

  // 在文档加载完成后初始化
  document.addEventListener("DOMContentLoaded", function () {
    // 创建右键菜单
    contextMenu = createContextMenu();

    // 点击其他地方隐藏菜单
    document.addEventListener("click", function (e) {
      if (!e.target.closest(".context-menu")) {
        contextMenu.classList.add("hidden");
      }
    });
  });
});
