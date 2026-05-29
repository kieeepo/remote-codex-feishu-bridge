# 安装手册

这份手册面向第一次安装的人。按步骤做完后，就可以用手机飞书远程控制本机 Codex。

## 一、准备条件

安装电脑建议使用 Windows 10 / Windows 11。

你需要提前准备：

- Node.js 20 或更高版本。
- OpenAI Codex 桌面端或 Codex CLI，并且已经登录。
- 飞书账号。
- 飞书开放平台的企业自建应用。
- 如果要生成或修改 Word / PPT，建议安装 Microsoft Office。

## 二、下载项目

把项目文件夹放到一个固定位置，例如：

```text
D:\RemoteCodexFeishuBridge
```

不要放到会频繁清理或同步冲突的位置。

## 三、安装依赖

双击：

```text
install.cmd
```

或者在项目目录打开 PowerShell：

```powershell
npm.cmd install
```

如果安装成功，项目目录里会出现 `node_modules`。

## 四、创建飞书自建应用

进入飞书开放平台，创建企业自建应用。

需要完成这些配置：

- 启用「机器人」能力。
- 在「权限管理」里开通这些权限：
  - `im:message`
  - `im:message:send_as_bot`
  - `im:resource`
  - `im:message:p2p_msg:readonly`
  - `im:message.group_at_msg:readonly`
- 在「事件与回调」里使用「长连接」。
- 订阅事件：
  - `im.message.receive_v1`
- 发布应用版本。
- 在飞书客户端里找到机器人，并给机器人发送 `测试`。

## 五、配置本地环境变量

复制 `.env.example` 为 `.env.local`。

如果你运行了 `install.cmd`，它会自动帮你创建 `.env.local`，但你仍然需要手动填写里面的值。

示例：

```env
FEISHU_APP_ID=cli_xxxxxxxxxxxxx
FEISHU_APP_SECRET=xxxxxxxxxxxxxxxx
FEISHU_ALLOWED_OPEN_IDS=ou_xxxxxxxxxxxxx
CODEX_BRIDGE_ENABLED=true
CODEX_BRIDGE_BIN=C:\Users\你的用户名\AppData\Local\OpenAI\Codex\bin\codex.exe
CODEX_BRIDGE_HOME=D:\CodexBridgeHome
CODEX_BRIDGE_SANDBOX=read-only
CODEX_BRIDGE_TIMEOUT_MS=300000
CLEANUP_RETENTION_DAYS=7
```

字段说明：

- `FEISHU_APP_ID`：飞书应用的 App ID。
- `FEISHU_APP_SECRET`：飞书应用的 App Secret。
- `FEISHU_ALLOWED_OPEN_IDS`：允许使用机器人的用户 open_id，建议只填你自己的。
- `CODEX_BRIDGE_BIN`：本机 `codex.exe` 路径。
- `CODEX_BRIDGE_HOME`：Codex Bridge 使用的独立配置目录，建议放在 D 盘。
- `CODEX_BRIDGE_SANDBOX`：默认保持 `read-only`。
- `CODEX_BRIDGE_TIMEOUT_MS`：单次任务最长等待时间，默认 5 分钟。
- `CLEANUP_RETENTION_DAYS`：自动清理几天前的临时文件。

## 六、配置 Codex 登录目录

如果 Codex Bridge 第一次运行时提示登录失败，可以把你当前 Codex 的登录配置复制到 `CODEX_BRIDGE_HOME`。

常见做法：

1. 新建目录，例如 `D:\CodexBridgeHome`。
2. 找到你当前 Codex 的配置目录。
3. 把 `auth.json` 和 `config.toml` 复制到 `D:\CodexBridgeHome`。
4. `.env.local` 里设置：

```env
CODEX_BRIDGE_HOME=D:\CodexBridgeHome
```

不要把这个目录发给别人。

## 七、启动机器人

双击：

```text
start-feishu.cmd
```

看到类似输出就说明启动成功：

```text
Remote Codex Feishu Agent starting...
Codex Bridge enabled.
Remote Codex Feishu Agent connected. Send '测试' to your bot.
```

然后在飞书里给机器人发：

```text
测试
```

如果机器人回复连接成功，说明安装完成。

## 八、怎么使用

直接聊天：

```text
帮我分析一下这个远程办公助手项目的亮点
```

生成 Word：

```text
帮我生成一个项目说明 Word 文档
```

生成 PPT：

```text
帮我做一个 6 页 PPT，主题是远程 Codex 办公助手
```

修改 Office 文件：

```text
先把 .docx 或 .pptx 发给机器人
再发：把第一页标题改成“项目背景”
```

图片识别：

```text
先把照片发给机器人
再发：识别这张图里的题目，并帮我写 Python 代码
```

## 九、常见问题

### npm 不能运行

PowerShell 如果拦截 `npm`，使用：

```powershell
npm.cmd install
```

### 机器人没有反应

检查：

- `start-feishu.cmd` 是否还开着。
- 飞书应用是否已经发布。
- 是否订阅了 `im.message.receive_v1`。
- 是否使用长连接。
- `.env.local` 里的 App ID / App Secret 是否正确。

### Codex Bridge 执行失败

检查：

- `CODEX_BRIDGE_BIN` 是否指向真实存在的 `codex.exe`。
- `CODEX_BRIDGE_HOME` 是否有有效登录配置。
- 本机 Codex 是否能正常运行。

### 生成文件太慢

PPT / Word 生成比普通聊天慢，尤其是 PowerPoint 自动化。建议把需求说清楚，减少反复修改。

## 十、安全建议

- 只把 `FEISHU_ALLOWED_OPEN_IDS` 设置为可信用户。
- 不要公开 `.env.local`。
- 不要公开 `CODEX_BRIDGE_HOME`。
- 不要把电脑暴露成公网服务，本项目默认使用飞书长连接，不需要开公网端口。
- 定期检查 `uploads` 和 `outputs`，项目会自动清理 7 天前临时文件。
