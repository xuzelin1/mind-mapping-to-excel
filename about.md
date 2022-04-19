## 技术背景

需要将一个 processOn 的思维导图转化到 Excel 表格中，方法有两个：第一，我们可以直接按照思维导图的结构，按照规则一个个填写到 Excel 中，合并单元格；第二，写个工具，也许以后还可以用呢。

## 思路与实现

### 技术选型和思路

技术上，用什么都可以，为了方便，我就选择用 node 进行开发。其中，因为需要写入到 Excel，因此需要借助一个 node 工具库 —— `node-xlsx`。

实现的思路是：
1. 通过接口先获取到数据，并且写入一个文件
2. 对数据进行格式转换，按照 `node-xlsx` 的数据格式要求，生成每一行的数据
3. 通过计算，确认合并的列规则
4. 生成 Excel 文件

### 分步实现

首先，拿到数据后，我们先观察 processOn 给出的数据格式，其实不难发现，processOn 的思维导图数据类似一棵树，其中，包括可能存在的 `leftChildren`。

```json
{
  "leftChildren": [],
  "children": [],
  "root": true,
  "theme": "delicate_dark",
  "id": "root",
  "title": "<font face=\"黑体\">抗疫囤货</font>",
  "structure": "mind_right"
}
```

知道了我们的源数据，我们接下来应该考虑的是，如何进行转换。如何进行转换则需要看看 `node-xlsx` 需要什么，也就是说，我们转换后的数据应该是怎样的。

```js
import xlsx from 'node-xlsx';
// Or var xlsx = require('node-xlsx').default;

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
var buffer = xlsx.build([{name: 'mySheetName', data: data}]); // Returns a buffer
```

可以看到，他要的数据其实非常简单，就是每一行的数据，组合成一个二维数组，那么非常简单，我们只需要对源数据进行深度遍历，拿到所有叶子节点的路径即可。

```js
/**
 * 获取树的叶子节点路径合集
 * @param {*} root - 根节点
 * @param {*} path - 路径
 * @param {*} level - 当前层
 */
function listPath(root, path){
    if((root.children || root.leftChildren).length === 0){// 叶子节点
        path = path + root.title;
        pathList.push(path.split('->').map((str) => {
            str = str.replace(/<[^>]+>/g,'');
            return str;
        })); // 将结果保存在list中
        return;
    }else{ // 非叶子节点
        path = path  + root.title + '->';
        // 子树
        const childs = root.children;
        // 左子树
        const leftChilds = root.leftChildren || [];
        
        //进行子节点的递归
        for(let i = 0; i < childs.length; i++){
            const childNode = childs[i];
            listPath(childNode, path);
        }

        // 存在左子树的情况
        if (leftChilds.length === 0) return;
        
        //进行子节点的递归
        for(let i = 0; i < leftChilds.length; i++){
            const childNode = leftChilds[i];
            listPath(childNode, path);
        }
    }
}

// 源数据
const data = fs.readFileSync('sourceData.JSON', 'utf8');
const sourceData = JSON.parse(data);

// 路径集合
const pathList = [];
if (sourceData) {
    listPath(sourceData, []);
}
```

到了这一步，我们基本上可以生成正确的 Excel 文件，而接下来要做的，就是合并单元格了。

```js
const range = {s: {c: 0, r: 0}, e: {c: 0, r: 3}}; // A1:A4
const sheetOptions = {'!merges': [range]};
```

结合 `node-xlsx` 的示例，可以知道 `range` 中的 `s` 代表的是开始合并的坐标，`e` 代表的是结束合并的坐标。上面的例子代表的就是从 `[0, 0]` 到 `[0, 3]` 合并，也就是 Excel 表格中的 `A1` 到 `A4`。

既然如此，我们再回到我们生成的 Excel 中看看:

![](https://raw.githubusercontent.com/xuzelin1/blog-img/main/img/202204182113930.png)

通过观察，我们知道第一列都是要合并的，因为第二列及后面的数据，都属于第一列的子集，也就是说，也就是说，从 `[0, 0]` 到 `[0, 75]`(`A1` 到 `A76`) 都是要合并的。再观察，你就不难发现，每一列，需要合并的单元格数量，都是以该点为根节点的树的叶子节点的数量。

为了知道在第几列，我们就需要知道当前在树的哪个层，因此引入一个 `level` 来记录。
而每一次开始的坐标，都是当前已完成路径遍历的节点数量。因此可以得出：

```js
function listPath(root, path, level = 0){
    // ...

    // 获取叶子节点的长度
    const childrenLen = getLeafNodeLen(root);

    // ...

    ranges.push({s: {c: level, r: pathList.length}, e: {c: level, r: pathList.length + childrenLen - 1}});

    // ...
}


```

而获取叶子节点的长度则可以通过广度优先搜索进行查找。

```js

/**
 * 获取叶子节点的长度
 */
function getLeafNodeLen(root){
    if (!root) return 0;
    if (root.children.length === 0 && (root.leftChildren || []).length === 0) return 1;
    else {
        let len = 0;
        for (let i = 0; i < root.children.length; i++) {
            len += getLeafNodeLen(root.children[i]);
        }
        for (let i = 0; i < (root.leftChildren || []).length; i++) {
            len += getLeafNodeLen(root.leftChildren[i]);
        }
        return len;
    }
}
```

## 总结

思维导图一般的实现方式是通过树来实现，而为了遍历树来获取想要的结果，往往可能需要通过深度优/广度优先的方式去实现。通过对数据的封装，我们就可以将数据转换，并最终生成我们想要的 Excel 文件。

- [Process On 示例思维导图](https://www.processon.com/view/link/625e1f931efad40734bf4d3e)