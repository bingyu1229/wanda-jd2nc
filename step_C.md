# 明细表生成：b.csv / GL_FREEVALUE.csv -> GL_DETAIL.csv

## 1. 运行方式

- 使用脚本：`test/step_C.ts`
- 在项目根目录执行：

```bash
npx tsx test/step_C.ts
```

- 默认输入：从当前执行目录开始，递归查找同时包含 `b.csv`、`GL_FREEVALUE.csv` 和 `pk.csv` 的目录。
- 调试单个目录时，可以传入目录路径：

```bash
npx tsx test/step_C.ts "test/015.001"
```

- 也可以直接传入 `b.csv` 路径：

```bash
npx tsx test/step_C.ts "test/015.001/b.csv"
```

- 默认输出：在输入目录内生成 `GL_DETAIL.csv`。

可选参数：

- `--utf8-bom`：输出带 UTF-8 BOM，便于 Windows Excel 打开

---

## 2. 生成 `GL_DETAIL.csv`

`GL_DETAIL.csv` 的数据行数与 `b.csv` 的数据行数保持一致。
`GL_DETAIL.csv` 会用 `b.csv` 的 E 列 `科目代码` 匹配 `pk.csv` 的 `SUBJCODE`，再取同一行的 `PK_ACCSUBJ` 填充明细表。`PK_CORP`、`PK_GLORG`、`PK_GLORGBOOK` 使用 `pk.csv` 第一条数据行中的固定值；`PK_GLBOOK` 固定为 `0001A9100000000JCNSC`。

`GL_DETAIL.csv` A 列 `ASSID` 取值规则：

- 若 `b.csv` 的 `科目名称` 中含有 `space-space`（脚本按 `-` 判断），则 `ASSID` 取同一条辅助核算数据在 `GL_FREEVALUE.csv` 中 I 列 `FREEVALUEID` 的值。
- 若 `科目名称` 中不含有 `space-space`，则 `ASSID` 留空。
- `step_C.ts` 会按辅助核算数据行顺序消费 `GL_FREEVALUE.csv`，确保例如第一条辅助核算明细对应 `1774A15010000000391F`。

字段规则：


| 字段                  | 规则                                                                     |
| ------------------- | ---------------------------------------------------------------------- |
| `ASSID`             | `GL_FREEVALUE.csv` 中对应行的 `FREEVALUEID`，或留空                             |
| `BANKACCOUNT`       | 留空                                                                     |
| `CHECKDATE`         | 留空                                                                     |
| `CHECKNO`           | 留空                                                                     |
| `CHECKSTYLE`        | 留空                                                                     |
| `CONTRASTFLAG`      | 留空                                                                     |
| `CONVERTFLAG`       | 留空                                                                     |
| `CREDITAMOUNT`      | 复制 `b.csv` 中 K 列 `贷方` 的值                                               |
| `CREDITQUANTITY`    | 固定 `0`                                                                 |
| `DEBITAMOUNT`       | 复制 `b.csv` 中 J 列 `借方` 的值                                               |
| `DEBITQUANTITY`     | 固定 `0`                                                                 |
| `DETAILINDEX`       | 复制 `b.csv` 中 AB 列 `分录号` 的值                                             |
| `DR`                | 固定 `0`                                                                 |
| `ERRMESSAGE`        | 留空                                                                     |
| `EXCRATE1`          | 固定 `0`                                                                 |
| `EXCRATE2`          | 固定 `1`                                                                 |
| `EXPLANATION`       | 复制 `b.csv` 中 D 列 `摘要` 的值                                               |
| `FRACCREDITAMOUNT`  | 固定 `0`                                                                 |
| `FRACDEBITAMOUNT`   | 固定 `0`                                                                 |
| `FREE1`             | 留空                                                                     |
| `FREE2`             | 留空                                                                     |
| `FREE3`             | 留空                                                                     |
| `FREE4`             | 留空                                                                     |
| `FREE5`             | 留空                                                                     |
| `LOCALCREDITAMOUNT` | 复制 `b.csv` 中 K 列 `贷方` 的值                                               |
| `LOCALDEBITAMOUNT`  | 复制 `b.csv` 中 J 列 `借方` 的值                                               |
| `MODIFYFLAG`        | 固定 `YYYYYYYYYYYYYYYY`                                                  |
| `OPPOSITESUBJ`      | 留空                                                                     |
| `PK_ACCSUBJ`        | 按 `b.csv` 科目代码匹配 `pk.csv.SUBJCODE` 后取 `PK_ACCSUBJ`                         |
| `PK_CORP`           | 取 `pk.csv` 第一条数据行的 `PK_CORP`                                           |
| `PK_CURRTYPE`       | 固定 `00010000000000000001`                                              |
| `PK_DETAIL`         | `1774A9` + 14 位 UUID；脚本从 `15020000000001` 开始递增                         |
| `PK_GLBOOK`         | 固定 `0001A9100000000JCNSC`                                              |
| `PK_GLORG`          | 取 `pk.csv` 第一条数据行的 `PK_GLORG`                                           |
| `PK_GLORGBOOK`      | 取 `pk.csv` 第一条数据行的 `PK_GLORGBOOK`                                       |
| `PK_INNERCORP`      | 留空                                                                     |
| `PK_INNERSOB`       | 留空                                                                     |
| `PK_SOB`            | 留空                                                                     |
| `PK_SOURCEPK`       | 留空                                                                     |
| `PK_VOUCHER`        | `0001DEFAULT` + 9 位 UUID；脚本从 `150000001` 开始递增                          |
| `PRICE`             | 固定 `0`                                                                 |
| `RECIEPTCLASS`      | 留空                                                                     |
| `TS`                | 固定 `2026-03-11 9:00:00`                                                |
| `DIRECTION`         | 若 `LOCALCREDITAMOUNT` 为 `0`，则为 `D`，否则为 `C`                             |
| `DISCARDFLAGV`      | 固定 `N`                                                                 |
| `ERRMESSAGE2`       | 留空                                                                     |
| `FREE6`             | 取 `b.csv` 中 B 列 `期间` 的月份。如 `2002.9`，则为 `09`                            |
| `NOV`               | 固定 `1`                                                                 |
| `PERIODV`           | 同 `FREE6`                                                              |
| `PK_MANAGERV`       | 留空                                                                     |
| `PK_SYSTEMV`        | 固定 `GL`                                                                |
| `PK_VOUCHERTYPEV`   | 固定 `0001DEFAULT000000001`                                              |
| `PREPAREDDATEV`     | 取 `b.csv` 中 A 列 `日期` 的值，并转为 `YYYY-MM-DD`，如 `2002/9/30` 转为 `2002-09-30` |
| `SIGNDATEV`         | 留空                                                                     |
| `VOUCHERKINDV`      | 留空                                                                     |
| `YEARV`             | 取 `b.csv` 中 B 列 `期间` 的年份。如 `2002.9`，则为 `2002`                          |
| `BUSIRECONNO`       | 留空                                                                     |
| `ERRMESSAGEH`       | 留空                                                                     |
| `FREE10`            | 留空                                                                     |
| `FREE11`            | 留空                                                                     |
| `FREE7`             | 留空                                                                     |
| `FREE8`             | 留空                                                                     |
| `FREE9`             | 留空                                                                     |
| `ISDIFFLAG`         | 固定 `N`                                                                 |
| `PK_OFFERDETAIL`    | 留空                                                                     |
| `PK_OTHERCORP`      | 留空                                                                     |
| `PK_OTHERORGBOOK`   | 留空                                                                     |
