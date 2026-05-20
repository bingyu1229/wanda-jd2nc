# 明细表生成：b.csv / GL_DETAIL.csv -> GL_VOUCHER.csv

## 1. 运行方式

- 使用脚本：`test/step_D.ts`
- 在项目根目录执行：

```bash
npx tsx test/step_D.ts
```

- 默认输入：从当前执行目录开始，递归查找同时包含 `b.csv`、`GL_DETAIL.csv`、`pk.csv` 和 `user.csv` 的目录。
- 调试单个目录时，可以传入目录路径：

```bash
npx tsx test/step_C.ts "test/015.001"
```

- 也可以直接传入 `b.csv` 路径：

```bash
npx tsx test/step_C.ts "test/015.001/b.csv"
```

- 默认输出：在输入目录内生成 `GL_VOUCHER.csv`。

可选参数：

- `--utf8-bom`：输出带 UTF-8 BOM，便于 Windows Excel 打开

---

## 2. 生成 `GL_VOUCHER.csv`

`GL_VOUCHER.csv` 的数据行数与 `GL_DETAIL.csv` 中`DETAILINDEX`为`1`的数据行数保持一致。

因为`GL_DETAIL.csv` 的数据行数与 `b.csv` 一致，且`GL_DETAIL.csv.DETAILINDEX`与`b.csv.分录号`保持一致，因此`GL_VOUCHER.csv可同时对照GL_DETAIL.csv`与`b.csv。`  

`GL_VOUCHER.csv` AQ 列 `TOTALCREDIT`和 AR 列 `TOTALDEBIT` 取值规则：

- 以`GL_DETAIL.csv.DETAILINDEX`=`1`开始，直到下一个`GL_DETAIL.csv.DETAILINDEX`=`1`之前的行数为一个取值区间，分别计算`GL_DETAIL.csv.CREDITAMOUNT`以及`GL_DETAIL.csv.DEBITAMOUNT`的求和结果，并填充到`TOTALCREDIT`和`TOTALDEBIT`中。

`GL_DETAIL.csv`：


| DETAILINDEX | CREDITAMOUNT | DEBITAMOUNT |
| ----------- | ------------ | ----------- |
| `1`         | `0.00`       | `60.00`     |
| `2`         | `0.00`       | `30.75`     |
| `3`         | `0.00`       | `10.00`     |
| `4`         | `80.00`      | `0.00`      |
| `1`         | `40.00`      | `0.00`      |
| `2`         | `0.00`       | `20.00`     |
| `1`         | `55.00`      | `0.00`      |
| `2`         | `0.00`       | `25.00`     |
| `3`         | `40.00`      | `0.00`      |


`GL_VOUCHER.csv`:


| TOTALCREDIT | TOTALDEBIT |
| ----------- | ---------- |
| `80.00`     | `100.00`   |
| `40.00`     | `20.00`    |
| `95.00`     | `26.00`    |


字段规则：


| 字段                  | 规则                                                                     |
| ------------------- | ---------------------------------------------------------------------- |
| `ADDCLASS`          | 留空                                                                     |
| `ATTACHMENT`        | 复制 `b.csv` 中 X 列 `附件` 的值                                               |
| `CHECKDATE`         | 留空                                                                     |
| `CONTRASTFLAG`      | 留空                                                                     |
| `CONVERTFLAG`       | 留空                                                                     |
| `DELETECLASS`       | 留空                                                                     |
| `DETAILMODFLAG`     | 固定 `Y`                                                                 |
| `DISCARDFLAG`       | 固定 `N`                                                                 |
| `DR`                | 固定 `0`                                                                 |
| `ERRMESSAGE`        | 留空                                                                     |
| `EXPLANATION`       | 留空                                                                     |
| `FREE1`             | 复制 `b.csv` 中 `PERIODV` 的值                                            |
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
| `PK_ACCSUBJ`        | 按 `b.csv` 科目代码匹配 `pk.csv.SUBJCODE` 后取 `PK_ACCSUBJ`                     |
| `PK_CORP`           | 取 `pk.csv` 第一条数据行的 `PK_CORP`                                           |
| `PK_CURRTYPE`       | 固定 `00010000000000000001`                                              |
| `PK_DETAIL`         | `1774A9` + 14 位 UUID；脚本从 `15020000000001` 开始递增                         |
| `PK_GLBOOK`         | 取 `pk.csv` 第一条数据行的 `PK_GLBOOK`；若无该列，则固定为 `0001A9100000000JCNSC`        |
| `PK_GLORG`          | 取 `pk.csv` 第一条数据行的 `PK_GLORG`                                          |
| `PK_GLORGBOOK`      | 取 `pk.csv` 第一条数据行的 `PK_GLORGBOOK`                                      |
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
| `PK_MANAGERV`       | `b.csv` N 列 `过账` 非空时，按 `user.csv.USER_NAME` 匹配后取 `CUSERID`；为空则留空       |
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


