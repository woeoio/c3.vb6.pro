# Project instructions

## CI

- 纯文档提交（只改 `ai/` 等说明性文件、对编译器产物零影响）推送时在提交信息里带 `[skip ci]`，不跑 GA。
- 动了编译器源码 / 夹具 / `emit-manifest.expected.txt` 的提交照旧跑，且**先推代码提交**、等 CI 起来再推文档提交——否则后一条的 `[skip ci]` 会让整次推送都不跑。
