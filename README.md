# 教案 Word 模板填充 OpenAPI 服务

这个目录是给 Dify OpenAPI / Swagger 插件用的云函数项目。部署后，Dify 可以调用接口，把 23 个教案字段填入 Word 模板，并生成 `.docx`。

## 1. 部署到 Vercel

1. 注册并登录 Vercel。
2. 新建项目，把本目录上传到 GitHub 后从 Vercel 导入；或者在本目录执行 `npx vercel`。
3. 部署成功后，拿到类似 `https://xxx.vercel.app` 的域名。
4. 浏览器打开 `https://xxx.vercel.app/api/health`，看到 `ok: true` 就说明接口可用。

## 2. 修改 OpenAPI Schema

打开 `openapi.yaml`，把：

```yaml
servers:
  - url: https://REPLACE_WITH_YOUR_VERCEL_DOMAIN.vercel.app
```

改成你的 Vercel 域名，例如：

```yaml
servers:
  - url: https://lesson-plan-docx-demo.vercel.app
```

然后在 Dify 的插件管理 / OpenAPI Schema 里导入这个 YAML。

## 3. Dify 工作流接线

推荐放置位置：

```text
开始
 -> 教师信息抽取
 -> 知识库检索案例与模板要求
 -> 生成教案字段
 -> 抽取 Word 模板字段
 -> OpenAPI 工具：generateLessonPlanDocx
 -> 结束
```

OpenAPI 工具节点配置：

| 参数 | 填法 |
|---|---|
| `filename` | 例如 `{{课程名称}}-第{{lesson_no}}次课教案.docx` |
| `data` | 传入“抽取 Word 模板字段”节点输出的 Object |

`data` 必须包含以下字段：

`lesson_no`, `hours`, `teaching_content`, `key_points`, `difficult_points`, `time_1_content`, `time_1_minutes`, `time_2_content`, `time_2_minutes`, `time_3_content`, `time_3_minutes`, `time_4_content`, `time_4_minutes`, `time_5_content`, `time_5_minutes`, `org_course_intro`, `org_lead_in`, `org_content_1`, `org_content_2`, `org_practice_interaction`, `org_summary`, `homework_report`, `teacher_signature_date`

## 4. 接口说明

主接口：

```http
POST /api/generate-lesson-plan-docx
Content-Type: application/json
```

返回：`application/vnd.openxmlformats-officedocument.wordprocessingml.document`

备用接口：

```http
POST /api/generate-lesson-plan-docx-json
Content-Type: application/json
```

返回：JSON，其中 `file_base64` 是 docx 文件内容。只有在 Dify 当前环境不能识别二进制文件返回时才用备用接口。

## 5. 为什么能保持模板格式

接口不会让大模型重新生成 Word，也不会用 Markdown 转 Word。它直接读取 `templates/lesson_plan_template.docx`，只替换 Word XML 里的 `{{字段名}}` 占位符，所以原模板里的表格、字体、页边距、嵌套表格和版式都会保留。
