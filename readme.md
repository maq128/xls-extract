# 项目说明

这是一个 Excel 数据处理工具，扫描所有 .xlsx 文件，按规定格式从里面查找并提取出数据，并把所有提取到的数据写到一个 .xlsx 文件里。

# 打包成 .exe

安装 `pkg` 打包工具：
```cmd
npm install -D pkg
```

基于 `package.json` 的配置信息进行打包：
```cmd
npx pkg .
```

## 打包成 32 位版本

最新版的 pkg 已经不支持 32 位打包，须安装旧版本。

支持 32 位打包的最后版本是 `pkg@4.5.1`，其内部依赖 `pkg-fetch@2.6.9`，打包版本支持到 `node14-win-x86`。

# 参考资料

[exceljs](https://github.com/exceljs/exceljs)

[nodejs 程序打包成 .exe](https://github.com/vercel/pkg)

[pkg-fetch 支持的 32 位打包版本](https://github.com/vercel/pkg-fetch/releases/tag/v2.6)
