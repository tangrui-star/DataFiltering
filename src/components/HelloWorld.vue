<template>
<div class="greetings">
    <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" />
    <button @click="processExcel">数据过滤</button>
    <button @click="downloadExcel" v-if="result.length">下载结果</button>
    <div v-if="result.length">
      <h3>处理结果:</h3>
      <ul>
        <li v-for="item in result" :key="item.商品 + item.分类">
          分类: {{ item.分类 }}, 商品: {{ item.商品 }}, 数量: {{ item.数量 }}
        </li>
      </ul>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue';
import * as XLSX from 'xlsx';

const result = ref([]);
let data = [];

const handleFileUpload = (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();
  
  reader.onload = (e) => {
    const binaryStr = e.target.result;
    const workbook = XLSX.read(binaryStr, { type: 'binary' });
    const sheet_name = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheet_name];

    // Convert the sheet to JSON format
    data = XLSX.utils.sheet_to_json(worksheet);
  };

  reader.readAsBinaryString(file);
};

const processExcel = () => {
  const productSales = {};

  // Aggregate quantities by product and category
  data.forEach((row) => {
    const product = row['商品'];
    const category = row['分类'];
    const quantity = row['数量'];

    const key = `${product} - ${category}`;
    if (!productSales[key]) {
      productSales[key] = 0;
    }
    productSales[key] += quantity;
  });

  // Convert the result into an array format
  const aggregatedResults = Object.keys(productSales).map((key) => {
    const [product, category] = key.split(' - ');
    return { 分类: category, 商品: product, 数量: productSales[key] };
  });

  // Sort results by category and quantity in descending order
  result.value = aggregatedResults.sort((a, b) => {
    if (a.分类 === b.分类) {
      return b.数量 - a.数量; // Sort by quantity in descending order
    }
    return a.分类.localeCompare(b.分类); // Sort by category
  });
};

const downloadExcel = () => {
  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(result.value);
  
  // Merges array to store merge information
  const merges = [];
  let startRow = 0;

  // Loop through the result to find adjacent identical entries in the 商品 column (A column)
  for (let i = 1; i <= result.value.length; i++) {
    // Check if current entry is different or if it's the end of the array
    if (i === result.value.length || result.value[i].分类 !== result.value[startRow].分类) {
      // If the range to merge is greater than 1
      if (i - startRow > 1) {
        merges.push({
          s: { r: startRow+1, c: 0 }, // Starting cell (row, column for A)
          e: { r: i, c: 0 } // Ending cell (row, column for A)
        });
      }
      // Update startRow to the current index
      startRow = i; 
    }
  }
    // Set cell styles for centering
    const cellStyle = {
    alignment: {
      horizontal: 'center',
      vertical: 'center',
    }
  };
  // console.log("sddfadf"+merges.toString);
  for(let sss=0;sss < merges.length-1;sss++){
    console.log("s:",merges[sss].s);
    console.log("e:",merges[sss].e);

  }

  // Append the sheet to the workbook
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Product Sales Summary');
  
  // Add merges to the sheet
  newSheet['!merges'] = merges;

  
  // Apply styles to each cell
  for (let i = 0; i < result.value.length; i++) {
    for (let key in result.value[i]) {
      const cellAddress = XLSX.utils.encode_cell({ r: i, c: 0 }); // Assuming you want to style the first column (A)
      if (!newSheet[cellAddress]) newSheet[cellAddress] = {}; // Create cell object if it doesn't exist
      newSheet[cellAddress].s = cellStyle; // Set style for the cell
    }
  }

  // Create a downloadable link
  const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  
  // Create a link element for download
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', 'product_sales_summary.xlsx');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
</script>

<style scoped></style>
