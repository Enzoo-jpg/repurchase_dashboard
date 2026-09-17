#!/usr/bin/env python3
# 构建脚本：把 mapping.js / pipeline.js / app.js 内联进 index.template.html。
#
# 用法：
#   python build.py            -> 生成 index.html（主程序内联，vendor 库保留外链，供 GitHub Pages / 本地目录部署）
#   python build.py --single   -> 生成 index.single.html（连 vendor 库也内联，约 2MB 真正单文件，
#                                 可直接发送给他人，双击即用、离线可用）
#
# 为什么要内联主程序：
#   1) 主程序完全自包含，无需额外请求；
#   2) 快照功能（doSnapshot）直接序列化当前页面生成自包含 HTML，
#      因此页面里的逻辑脚本必须是内联的，否则下载/双击打开的快照会因
#      缺少外部 mapping.js/pipeline.js/app.js 而无法运行。
#   vendor/xlsx.full.min.js / exceljs.min.js 默认保留为外链（在线加载即可；
#   快照为只读视图不会用到它们，生成时会被剥离）。--single 模式下两个库
#   以内联 id 标记块形式嵌入，app.js 的快照剥离逻辑同样能识别并移除。
import argparse
import io
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def read(name):
    with io.open(os.path.join(HERE, name), encoding="utf-8") as f:
        return f.read()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--single", action="store_true",
                        help="生成单文件 index.single.html（内联 vendor 库）")
    args = parser.parse_args()

    template = read("index.template.html")
    mapping = read("mapping.js")
    pipeline = read("pipeline.js")
    app = read("app.js")

    html = template
    html = html.replace(
        '<script src="mapping.js"></script>',
        "<script>\n" + mapping + "\n</script>",
    )
    html = html.replace(
        '<script src="pipeline.js"></script>',
        "<script>\n" + pipeline + "\n</script>",
    )
    html = html.replace(
        '<script src="app.js"></script>',
        "<script>\n" + app + "\n</script>",
    )

    if args.single:
        # 内联 vendor 库（已验证两个 min 文件均不含 </script> 字面量，内联安全）
        xlsx = read(os.path.join("vendor", "xlsx.full.min.js"))
        exceljs = read(os.path.join("vendor", "exceljs.min.js"))
        html = html.replace(
            '<script src="vendor/xlsx.full.min.js"></script>',
            "<script id=\"__vnd_xlsx__\">\n" + xlsx + "\n</script>",
        )
        html = html.replace(
            '<script src="vendor/exceljs.min.js"></script>',
            "<script id=\"__vnd_exceljs__\">\n" + exceljs + "\n</script>",
        )
        out = os.path.join(HERE, "index.single.html")
    else:
        out = os.path.join(HERE, "index.html")

    with io.open(out, "w", encoding="utf-8") as f:
        f.write(html)

    print("build ok ->", out, "(%d bytes)" % len(html))


if __name__ == "__main__":
    main()
