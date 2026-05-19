# 辅助核算表生成：HTM -> b.csv / GL_FREEVALUE.csv

## 1. 运行方式

- 使用脚本：`test/step_B.ts`
- 在项目根目录执行：

```bash
npx tsx test/step_B.ts
```

- 默认输入：从当前执行目录开始，递归查找所有子目录中的 `会计分类序时簿.htm`。
- 调试单个文件时，也可以显式传入输入路径：

```bash
npx tsx test/step_B.ts "test/015.001/会计分类序时簿.htm"
```

- 默认输出：每个输入文件所在目录内生成：
  - `b.csv`
  - `GL_FREEVALUE.csv`

可选参数：

- `--encoding gb18030`：输入编码，默认 `gb18030`，兼容 `gb2312`
- `--utf8-bom`：输出带 UTF-8 BOM，便于 Windows Excel 打开
- `--excel-text-cols 0,1`：指定列以 Excel 文本公式形式写出
- `--no-normalize-col-e`：跳过列 E `科目代码` 规范化

说明：转换阶段不按数字解析单元格，避免日期、期间、编码等文本被截断或丢失末尾 `0`。

---

## 2. 生成 `b.csv`

### 2.1 HTM 基础转换

- 跳过 HTML 表内第 1 行占位表头：`第1列` 到 `第27列`。
- 保留业务表头作为 CSV 首行：`日期`、`期间`、`凭证号`、`摘要`、`科目代码` 等。
- 在最后新增第 28 列，表头为 `分录号`。

### 2.2 规范化列 E：`科目代码`

仅处理“整格由英文字母、数字与点号组成”的科目代码；表头、中文、其他非代码文本保持原样。

与 `a.csv` 中的 `原科目代码` 匹配，返回 `a.csv` 中对应的 A 列 `科目代码` 的值，并替换现在列 E `科目代码` 中的值。

### 2.3 向下填充空值

以下列如果为空，则使用上一条非空值填充：

- A 列：`日期`
- B 列：`期间`
- C 列：`凭证号`
- V 列：`业务日期`

### 2.4 填充 `分录号`

新增列 `分录号` 从 `1` 开始递增。

递增分组依据为：

- `日期`
- `期间`
- `凭证号`

当上述三列组成的新凭证键变化时，`分录号` 重新从 `1` 开始。

### 2.5 清理金额列

替换以下列中的 `'0` 为 `0`，并保留两位小数：

- I 列：`原币金额`
- J 列：`借方`
- K 列：`贷方`

### 2.6 清理列 F：`科目名称`

删除 `科目名称` 中的：

- 英文小写 `b`
- 英文逗号 `,`

注意：不要删除中文逗号 `，`。

---

## 3. 生成 `GL_FREEVALUE.csv`

`GL_FREEVALUE.csv` 由 `b.csv` 的 F 列 `科目名称` 中的辅助核算值生成。

生成规则：

1. 扫描 `b.csv` 数据行的 `科目名称`。
2. 若 `科目名称` 中含有 `space-space`（脚本按 `-` 判断），则取第一个 `space-space` 之后的文本作为辅助核算值。
  - 示例：`其他应收款_备用金 - 部门:ZJB - 总经办/职员:HMC - 何明春`
  - 辅助核算值：`部门:ZJB - 总经办/职员:HMC - 何明春`
3. 如果 `科目名称` 不包含 `space-space`，则不写入 `GL_FREEVALUE.csv`。
4. 不需要去重；每条辅助核算数据行生成一条 `GL_FREEVALUE.csv` 记录。

输出表头固定为：

```text
ASSINDEX,CHECKCOUNT,CHECKTYPE,CHECKVALUE,DR,FREE1,FREE2,FREE3,FREEVALUEID,PK_FREEVALUE,TS,VALUECODE,VALUENAME
```

字段规则：


| 字段                          | 规则                                                     |
| --------------------------- | ------------------------------------------------------ |
| `ASSINDEX`                  | 固定 `0`                                                 |
| `CHECKCOUNT`                | 固定 `1`                                                 |
| `CHECKTYPE`                 | 固定 `0001A9100000000JCKUS`                              |
| `CHECKVALUE`                | `0001A92JDT` + 9 位 UUID + `U`；UUID 从 `150000001` 开始递增  |
| `DR`                        | 固定 `0`                                                 |
| `FREE1` / `FREE2` / `FREE3` | 留空                                                     |
| `FREEVALUEID`               | `1774A` + 14 位 UUID + `F`；UUID 从 `15010000000391` 开始递增 |
| `PK_FREEVALUE`              | `1774A` + 14 位 UUID + `P`；UUID 与 `FREEVALUEID` 相同      |
| `TS`                        | 固定 `2026/3/5 16:26`                                    |
| `VALUECODE`                 | 五位序号，从 `00001` 开始                                      |
| `VALUENAME`                 | 辅助核算值原文                                                |


示例：

```csv
ASSINDEX,CHECKCOUNT,CHECKTYPE,CHECKVALUE,DR,FREE1,FREE2,FREE3,FREEVALUEID,PK_FREEVALUE,TS,VALUECODE,VALUENAME
0,1,0001A9100000000JCKUS,0001A92JDT150000001U,0,,,,1774A15010000000391F,1774A15010000000391P,2026/3/5 16:26,00001,部门:ZJB - 总经办/职员:HMC - 何明春
```

---

## 4. 生成 `name_sql.txt`

`name_sql.txt` 由 `b.csv` 的 L、M、N 列生成：

- L 列：`制单`
- M 列：`审核`
- N 列：`过账`

生成规则：

1. 扫描三列中的所有姓名。
2. 去除空值，并按首次出现顺序去重。
3. `USER_NAME` 使用原中文姓名。
4. `USER_CODE` 使用姓名拼音小写并追加 `(jindie)`。

输出分为两段：

```text
USER_CODE
(
'yinshijun(jindie)',
'wangxingquan(jindie)'
)

USER_NAME
(
'尹世军',
'王兴全'
)
```

