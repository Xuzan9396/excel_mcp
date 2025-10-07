#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

// 获取平台和架构信息
const platform = process.platform;
const arch = process.arch;

console.log('');
console.log('╔════════════════════════════════════════════════╗');
console.log('║       🚀 Excel MCP Server 安装程序       ║');
console.log('╚════════════════════════════════════════════════╝');
console.log('');
console.log(`   平台: ${platform}`);
console.log(`   架构: ${arch}`);

// 映射到二进制文件名
function getBinaryName() {
  let platformName = '';
  let archName = '';
  let extension = '';

  // 平台映射
  switch (platform) {
    case 'darwin':
      platformName = 'darwin';
      break;
    case 'linux':
      platformName = 'linux';
      break;
    case 'win32':
      platformName = 'windows';
      extension = '.exe';
      break;
    default:
      throw new Error(`不支持的平台: ${platform}`);
  }

  // 架构映射
  switch (arch) {
    case 'x64':
      archName = 'amd64';
      break;
    case 'arm64':
      archName = 'arm64';
      break;
    default:
      throw new Error(`不支持的架构: ${arch}`);
  }

  return `excel-mcp-${platformName}-${archName}${extension}`;
}

// 获取最新版本号
function getLatestVersion() {
  const packageJson = require('./package.json');
  return packageJson.version;
}

// 延迟函数
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 下载文件（单次尝试）
function downloadFileOnce(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    let timeoutId;

    const request = https.get(url, { timeout: 30000 }, (response) => {
      // 清除超时定时器
      if (timeoutId) clearTimeout(timeoutId);

      // 处理重定向
      if (response.statusCode === 302 || response.statusCode === 301) {
        file.close();
        fs.unlink(dest, () => {});
        return downloadFileOnce(response.headers.location, dest)
          .then(resolve)
          .catch(reject);
      }

      if (response.statusCode !== 200) {
        file.close();
        fs.unlink(dest, () => {});
        reject(new Error(`HTTP ${response.statusCode}`));
        return;
      }

      const totalSize = parseInt(response.headers['content-length'], 10);
      let downloadedSize = 0;
      let lastUpdate = Date.now();
      const startTime = Date.now();

      response.on('data', (chunk) => {
        downloadedSize += chunk.length;
        const now = Date.now();

        // 每 100ms 更新一次进度条
        if (now - lastUpdate > 100) {
          const percent = Math.floor((downloadedSize / totalSize) * 100);
          const downloaded = (downloadedSize / 1024 / 1024).toFixed(2);
          const total = (totalSize / 1024 / 1024).toFixed(2);

          // 计算速度
          const elapsed = (now - startTime) / 1000;
          const speed = downloadedSize / elapsed / 1024 / 1024; // MB/s

          // 计算剩余时间
          const remaining = (totalSize - downloadedSize) / (downloadedSize / elapsed);
          const eta = remaining > 0 && remaining < 3600 ? `${Math.ceil(remaining)}s` : '--';

          // 绘制进度条 [=====>    ]
          const barLength = 20;
          const filled = Math.floor(barLength * percent / 100);
          const bar = '='.repeat(filled) + '>'.padEnd(barLength - filled);

          process.stdout.write(`\r   [${bar}] ${percent}% | ${downloaded}/${total}MB | ${speed.toFixed(2)}MB/s | ETA: ${eta}`);
          lastUpdate = now;
        }
      });

      response.pipe(file);

      file.on('finish', () => {
        file.close();
        if (timeoutId) clearTimeout(timeoutId);
        resolve();
      });

      file.on('error', (err) => {
        if (timeoutId) clearTimeout(timeoutId);
        fs.unlink(dest, () => {});
        reject(err);
      });
    });

    // 设置单次下载超时（30秒）
    timeoutId = setTimeout(() => {
      request.destroy();
      file.close();
      fs.unlink(dest, () => {});
      reject(new Error('单次下载超时'));
    }, 30000);

    request.on('error', (err) => {
      if (timeoutId) clearTimeout(timeoutId);
      file.close();
      fs.unlink(dest, () => {});
      reject(err);
    });

    request.on('timeout', () => {
      request.destroy();
      file.close();
      fs.unlink(dest, () => {});
      reject(new Error('连接超时'));
    });
  });
}

// 下载文件（带有限重试，避免无限阻塞导致 npx 缓存不能清理）
async function downloadFile(url, dest) {
  let attempt = 0;
  const maxRetryDelay = 30000; // 最大重试间隔 30 秒
  const maxAttempts = 6; // 最多重试 6 次（约 1+2+4+8+16+30 ≈ 61 秒）

  console.log(`📥 下载: ${url}`);

  while (attempt < maxAttempts) {
    attempt++;

    try {
      await downloadFileOnce(url, dest);
      console.log('\n✓ 下载完成');
      return;
    } catch (err) {
      // 计算重试间隔：指数退避，最大 30 秒
      const retryDelay = Math.min(1000 * Math.pow(2, attempt - 1), maxRetryDelay);

      console.log(`\n⚠️  下载失败 (尝试 ${attempt}/${maxAttempts}): ${err.message}`);
      if (attempt >= maxAttempts) {
        throw new Error(`多次重试仍失败（共 ${maxAttempts} 次），请稍后再试。`);
      }
      console.log(`   ${retryDelay / 1000} 秒后重试...`);

      await delay(retryDelay);

      // 清理进度输出
      process.stdout.write('\r\x1b[K');
    }
  }
}

// 清理 npm 在更新时创建但未删除的备份目录（.excel-mcp-*）
function cleanupBackupDirs() {
  try {
    // 当前目录: .../node_modules/@xuzan/excel-mcp
    const vendorDir = path.dirname(__dirname); // .../node_modules/@xuzan
    const entries = fs.readdirSync(vendorDir, { withFileTypes: true });
    const backups = entries
      .filter((e) => e.isDirectory() && e.name.startsWith('.excel-mcp-'))
      .map((e) => path.join(vendorDir, e.name));

    for (const dir of backups) {
      try {
        // 仅删除我们自己的备份目录名称前缀，避免误删
        fs.rmSync(dir, { recursive: true, force: true });
        console.log(`🧹 已清理备份目录: ${path.basename(dir)}`);
      } catch (e) {
        // 忽略单个清理失败
        console.warn(`⚠️  清理失败: ${path.basename(dir)} - ${e.message}`);
      }
    }
  } catch (e) {
    // 忽略读取失败
  }
}

// 主安装流程
async function install() {
  try {
    const binaryName = getBinaryName();
    const version = getLatestVersion();

    console.log('');
    console.log(`   📦 版本: v${version}`);
    console.log(`   📄 文件: ${binaryName}`);
    console.log('');

    // 下载 URL
    const downloadUrl = `https://github.com/Xuzan9396/excel_mcp/releases/download/v${version}/${binaryName}`;

    // 目标路径
    const binDir = path.join(__dirname, 'bin');
    const targetName = platform === 'win32' ? 'excel-mcp.exe' : 'excel-mcp';
    const binaryPath = path.join(binDir, targetName);

    // 创建 bin 目录
    if (!fs.existsSync(binDir)) {
      fs.mkdirSync(binDir, { recursive: true });
    }

    // 下载二进制文件（带无限重试）
    await downloadFile(downloadUrl, binaryPath);

    // 设置执行权限（Unix 系统）
    if (platform !== 'win32') {
      fs.chmodSync(binaryPath, 0o755);
      console.log('✓ 设置执行权限');
    }

    // 验证二进制文件
    console.log('🔍 验证二进制文件...');
    const stats = fs.statSync(binaryPath);
    console.log(`   大小: ${(stats.size / 1024 / 1024).toFixed(2)} MB`);

    console.log('');
    console.log('╔════════════════════════════════════════════════╗');
    console.log('║           ✅ 安装成功！                  ║');
    console.log('╚════════════════════════════════════════════════╝');
    console.log('');
    console.log('💡 使用方法:');
    console.log('   npx @xuzan/excel-mcp');
    console.log('');
    console.log('📖 完整文档:');
    console.log('   https://github.com/Xuzan9396/excel_mcp');
    console.log('');

    // 清理可能遗留的备份目录，避免下次 npx 更新时报 ENOTEMPTY
    cleanupBackupDirs();

  } catch (error) {
    console.error('\n❌ 安装失败:', error.message);
    console.error('\n💡 你可以手动下载二进制文件:');
    console.error(`   https://github.com/Xuzan9396/excel_mcp/releases\n`);
    process.exit(1);
  }
}

// 运行安装
install();
