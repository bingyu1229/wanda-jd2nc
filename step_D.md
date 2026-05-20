# 凭证表生成：b.csv / GL_DETAIL.csv -> GL_VOUCHER.csv

## 1. 运行方式

- 使用脚本：`test/step_D.ts`（封装调用 `test/script_D.ts`）
- 在项目根目录执行：

```bash
npx tsx test/step_D.ts
```

- 默认输入：从当前执行目录开始，递归查找同时包含 `b.csv`、`GL_DETAIL.csv`、`pk.csv` 和 `user.csv` 的目录。
- 调试单个目录时，可以传入目录路径：

```bash
npx tsx test/step_D.ts "test/015.001"
```

- 也可以直接传入 `b.csv` 路径：

```bash
npx tsx test/step_D.ts "test/015.001/b.csv"
```

- 默认输出：在输入目录内生成 `GL_VOUCHER.csv`。

可选参数：

- `--utf8-bom`：输出带 UTF-8 BOM，便于 Windows Excel 打开

---

## 2. 生成 `GL_VOUCHER.csv`

`GL_VOUCHER.csv` 的数据行数与 `GL_DETAIL.csv` 中`DETAILINDEX`为`1`的数据行数保持一致。

因为 `GL_DETAIL.csv` 的数据行数与 `b.csv` 一致，且 `GL_DETAIL.csv.DETAILINDEX` 与 `b.csv.分录号` 保持一致，因此 `GL_VOUCHER.csv` 可同时对照 `GL_DETAIL.csv` 与 `b.csv`。  

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
| `95.00`     | `25.00`    |


字段规则：


| 字段                                                | 规则                                                               |
| ------------------------------------------------- | ---------------------------------------------------------------- |
| `ADDCLASS`                                        | 留空                                                               |
| `ATTACHMENT`                                      | 复制 `b.csv` 中 X 列 `附件` 的值                                         |
| `CHECKEDDATE`                                     | 留空                                                               |
| `CONTRASTFLAG`                                    | 留空                                                               |
| `CONVERTFLAG`                                     | 留空                                                               |
| `DELETECLASS`                                     | 留空                                                               |
| `DETAILMODFLAG`                                   | 固定 `Y`                                                           |
| `DISCARDFLAG`                                     | 固定 `N`                                                           |
| `DR`                                              | 固定 `0`                                                           |
| `ERRMESSAGE`                                      | 留空                                                               |
| `EXPLANATION`                                     | 留空                                                               |
| `FREE1`                                           | 复制 `GL_DETAIL.csv` 中 `PERIODV` 的值                                |
| `FREE10`                                          | 固定 `VOUCHERNEWADD`                                               |
| `FREE2/FREE3/FREE4/FREE5/FREE6/FREE7/FREE8/FREE9` | 留空                                                               |
| `MODIFYCLASS`                                     | 留空                                                               |
| `MODIFYFLAG`                                      | 固定 `YYY`                                                         |
| `NO`                                              | 固定 `1`                                                           |
| `PERIOD`                                          | 复制 `GL_DETAIL.csv` 中 `PERIODV` 的值                                |
| `PK_CASHER`                                       | 留空                                                               |
| `PK_CHECKED`                                      | `b.csv` M 列 `审核` 非空时，按 `user.csv.USER_NAME` 匹配后取 `CUSERID`；为空则留空 |
| `PK_CORP`                                         | 复制 `GL_DETAIL.csv` 中 `PK_CORP`的值                                 |
| `PK_GLBOOK`                                       | 复制 `GL_DETAIL.csv` 中 `PK_GLBOOK`的值                               |
| `PK_GLORG`                                        | 复制 `GL_DETAIL.csv` 中 `PK_GLORG`的值                                |
| `PK_GLORGBOOK`                                    | 复制 `GL_DETAIL.csv` 中 `PK_GLORGBOOK`的值                            |
| `PK_MANAGER`                                      | `b.csv` N 列 `过账` 非空时，按 `user.csv.USER_NAME` 匹配后取 `CUSERID`；为空则留空 |
| `PK_PREPARED`                                     | `b.csv` L 列 `制单` 非空时，按 `user.csv.USER_NAME` 匹配后取 `CUSERID`；为空则留空 |
| `PK_SOB`                                          | 留空                                                               |
| `PK_SOURCEPK`                                     | 留空                                                               |
| `PK_SYSTEM`                                       | 固定 `GL`                                                          |
| `PK_VOUCHER`                                      | 复制 `GL_DETAIL.csv` 中 `PK_VOUCHER`的值                              |
| `PK_VOUCHERTYPE`                                  | 固定 `0001DEFAULT000000001`                                        |
| `PREPAREDDATE`                                    | 复制 `b.csv` 中 A列`日期`的值                                            |
| `SIGNDATE`                                        | 留空                                                               |
| `SIGNFLAG`                                        | 固定 `Y`                                                           |
| `TALLYDATE`                                       | 同`PREPAREDDATE`                                                  |
| `TOTALCREDIT`                                     | 期间内 `GL_DETAIL.csv` 中 `CREDITAMOUNT` 的求和结果                       |
| `TOTALDEBIT`                                      | 期间内 `GL_DETAIL.csv` 中 `DEBITAMOUNT` 的求和结果                        |
| `TS`                                              | 固定 `2026-03-11 9:00:00`                                          |
| `VOUCHERKIND`                                     | 固定 `0`                                                           |
| `YEAR`                                            | 复制 `GL_DETAIL.csv` 中 `YEARV`的值                                   |
| `ERRMESSAGEH`                                     | 留空                                                               |
| `ISDIFFLAG`                                       | 固定 `N`                                                           |
| `OFFERVOUCHER`                                    | 留空                                                               |
