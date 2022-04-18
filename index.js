const xlsx = require('node-xlsx');
const fs = require('fs');

const sourceData = JSON.parse(`

`)

const pathList = [];
const ranges = [];
listPath(sourceData, []);

const range = {s: {c: 0, r: 0}, e: {c: 0, r: pathList.length - 1}};
const sheetOptions = {'!merges': [range, ...ranges]};

const name = sourceData.title;
const buffer = xlsx.build([{name, data: pathList}], {sheetOptions}); // Returns a buffer

fs.writeFile('processon.xlsx', buffer, function(err) {
    if (err) {
        console.log("Write failed: " + err);
        return;
    }
});

/**
 * 获取树的叶子节点路径合集
 * @param {*} root - 根节点
 * @param {*} path - 路径
 * @param {*} level - 当前层
 */
function listPath(root, path, level = 0){

    if(root.children.length === 0){// 叶子节点
        path = path + root.title;
        pathList.push(path.split('->')); // 将结果保存在list中
        return;
    }else{ // 非叶子节点

        path = path  + root.title + '->';

        // 获取叶子节点的长度
        const childs = root.children;
        const childrenLen = getLeafNodeLen(root);
        ranges.push({s: {c: level, r: pathList.length}, e: {c: level, r: pathList.length + childrenLen - 1}});
        
        //进行子节点的递归
        for(let i = 0; i < childs.length; i++){
            const childNode = childs[i];
            listPath(childNode, path, level + 1);
        }
    }
}

/**
 * 获取叶子节点的长度
 */
function getLeafNodeLen(root){
    if (!root) return 0;
    if (root.children.length === 0) return 1;
    else {
        let len = 0;
        for (let i = 0; i < root.children.length; i++) {
            len += getLeafNodeLen(root.children[i]);
        }
        return len;
    }
}
