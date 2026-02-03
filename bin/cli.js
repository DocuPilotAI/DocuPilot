#!/usr/bin/env node

import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const projectRoot = join(__dirname, '..');

const args = process.argv.slice(2);
const command = args[0] || 'start';

function printHelp() {
  console.log(`
DocuPilot - 智能 Office 助手

用法:
  docupilot <command> [options]

命令:
  start       启动开发服务器 (HTTPS)
  dev         启动开发服务器 (HTTP)
  build       构建生产版本
  help        显示帮助信息

选项:
  --port, -p  指定端口号 (默认: 3000)
  --host, -h  指定主机地址 (默认: localhost)

示例:
  docupilot start           # 启动 HTTPS 开发服务器
  docupilot start -p 3001   # 在端口 3001 启动
  docupilot build           # 构建生产版本

注意:
  Office Add-in 需要 HTTPS 连接。
  首次运行时会自动生成自签名证书。
`);
}

function runCommand(cmd, cmdArgs = []) {
  console.log(`\n🚀 正在启动 DocuPilot...\n`);
  
  const child = spawn('npm', ['run', cmd, ...cmdArgs], {
    cwd: projectRoot,
    stdio: 'inherit',
    shell: true,
  });

  child.on('error', (error) => {
    console.error(`启动失败: ${error.message}`);
    process.exit(1);
  });

  child.on('close', (code) => {
    process.exit(code);
  });
}

switch (command) {
  case 'start':
    runCommand('dev:https');
    break;
  
  case 'dev':
    runCommand('dev');
    break;
  
  case 'build':
    runCommand('build');
    break;
  
  case 'help':
  case '--help':
  case '-h':
    printHelp();
    break;
  
  default:
    console.error(`未知命令: ${command}`);
    printHelp();
    process.exit(1);
}
