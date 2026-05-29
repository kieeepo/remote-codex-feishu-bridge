# Remote Codex Feishu Bridge

A local Feishu-to-Codex bridge for remote phone control and document generation.

这是一个本地运行的飞书到 Codex 桥接工具，可以让用户通过手机飞书机器人，把任务发送到电脑上的 Codex 执行，并返回文本、Word、PPT 等结果。

## Features

- Send Feishu messages to local Codex
- Generate Word and PowerPoint files
- Upload Office files and return modified copies
- Restrict access with Feishu open_id allowlist
- Run locally without exposing a public network port

## Quick Start

1. Download `remote-codex-feishu-bridge.zip` from Releases.
2. Extract it to a fixed folder.
3. Run `install.cmd`.
4. Copy `.env.example` to `.env.local`.
5. Fill in your Feishu app credentials and Codex path.
6. Run `start-feishu.cmd`.

## Security

- Do not publish `.env.local`.
- Do not publish Feishu App Secret.
- Only allow trusted Feishu open_ids.
- Keep Codex sandbox as `read-only` unless file generation is needed.
- Do not expose this tool as a public web service.

## License

MIT
