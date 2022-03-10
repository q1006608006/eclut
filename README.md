# eclut

excel数据合并工具

* [eclut](#eclut)
    * [主界面截图：](#主界面截图)
  * [选择模板：](#选择模板)
    * [示例1：](#示例1)
    * [示例2：](#示例2)
  * [读取数据：](#读取数据)
    * [对应示例1：](#对应示例1)
    * [对应示例2：](#对应示例2)
  * [合并数据](#合并数据)
    * [示例3：](#示例3)
  * [导出数据：](#导出数据)
    * [示例：](#示例)


### 主界面截图：
![image](https://user-images.githubusercontent.com/31004882/157615014-f2e39e60-1ae8-45fe-aac2-180909014a79.png)

## 选择模板：
模板仅支持xls文件，支持单个文件多个sheet（sheet名称必须一致）

### 示例1：
`{name}`表示读取当前单元格的值

![image](https://user-images.githubusercontent.com/31004882/157615717-efaf9f6e-2e45-462b-b15e-66f10d1f6645.png)
![image](https://user-images.githubusercontent.com/31004882/157615940-57742e49-d09e-4440-bcf9-69c2a200135b.png)

### 示例2：
`[name]`表示读取该列单元格（从上往下）
`[tel:1]`表示读取该列单元格（储存结果将保存到tel列）,**1**表示最多向下读取**1**个单元格

![image](https://user-images.githubusercontent.com/31004882/157615637-be27a26b-b32c-428a-ba5b-68665d553ffe.png)


## 读取数据：
选择与模板对应的excel文件，支持xls及xlsx格式，可以多次读取，也可以一次选择多个文件（需同一格式文件）

### 对应示例1：
![image](https://user-images.githubusercontent.com/31004882/157617243-7255eff9-e0a2-43b6-9f9a-9af4945640a7.png)
![image](https://user-images.githubusercontent.com/31004882/157617269-37c728fc-3ece-4ed8-a154-3cf8d7764d21.png)
![image](https://user-images.githubusercontent.com/31004882/157617119-aaa78c90-a23c-462f-8214-40546cd83c16.png)

**可使用sheet关联值关联数据** （所有列表长度必须一致）
![image](https://user-images.githubusercontent.com/31004882/157617574-66cbf3d0-ea0a-4e36-8f08-9659d23f6f89.png)

### 对应示例2：

![image](https://user-images.githubusercontent.com/31004882/157617189-a384f5c2-4106-4d7a-80c4-da9a320461a9.png)
![image](https://user-images.githubusercontent.com/31004882/157618134-4f1c5077-becc-4046-aa24-330a5d642e08.png)

## 合并数据
在已有数据的情况下，不清空数据，重新选择模板，读取对应文件即可

### 示例3：
![image](https://user-images.githubusercontent.com/31004882/157618408-343bc21d-91bb-468d-88a5-b9b49f7faed8.png)

**可使用`关联值`合并数据**
![image](https://user-images.githubusercontent.com/31004882/157618730-f2fe88ca-b3b5-4663-a317-20b187c5f964.png)

## 导出数据：
选择模板，导出数据

### 示例：
以示例3合并数据为例
![image](https://user-images.githubusercontent.com/31004882/157619579-addced4c-55e4-4aa5-b688-fba906828736.png)
![image](https://user-images.githubusercontent.com/31004882/157619677-5e63c275-b8a7-4aac-8dad-ace75d6fb65d.png)

