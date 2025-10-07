#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

// 获取平台和架构信息
const platform = process.platform;
const arch = process.arch;

console.log('🚀 Excel MCP 安装中...');
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
      let lastPercent = 0;

      response.on('data', (chunk) => {
        downloadedSize += chunk.length;
        const percent = Math.floor((downloadedSize / totalSize) * 100);
        if (percent > lastPercent && percent % 10 === 0) {
          process.stdout.write(`\r   进度: ${percent}%`);
          lastPercent = percent;
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

// 下载文件（带无限重试）
async function downloadFile(url, dest) {
  let attempt = 0;
  const maxRetryDelay = 30000; // 最大重试间隔 30 秒

  console.log(`📥 下载: ${url}`);

  while (true) {
    attempt++;

    try {
      await downloadFileOnce(url, dest);
      console.log('\n✓ 下载完成');
      return;
    } catch (err) {
      // 计算重试间隔：指数退避，最大 30 秒
      // 第1次: 1s, 第2次: 2s, 第3次: 4s, 第4次: 8s, 第5次: 16s, 第6次+: 30s
      const retryDelay = Math.min(1000 * Math.pow(2, attempt - 1), maxRetryDelay);

      console.log(`\n⚠️  下载失败 (尝试 ${attempt}): ${err.message}`);
      console.log(`   ${retryDelay / 1000} 秒后重试...`);

      await delay(retryDelay);

      // 清理进度输出
      process.stdout.write('\r\x1b[K');
    }
  }
}

// 主安装流程
async function install() {
  try {
    const binaryName = getBinaryName();
    const version = getLatestVersion();

    console.log(`   版本: v${version}`);
    console.log(`   二进制: ${binaryName}\n`);

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

    // 下载二进制文件
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

    console.log('\n✅ 安装成功！');
    console.log('\n💡 使用方法:');
    console.log('   npx @xuzan/excel-mcp');
    console.log('\n📖 更多信息:');
    console.log('   https://github.com/Xuzan9396/excel_mcp\n');

  } catch (error) {
    console.error('\n❌ 安装失败:', error.message);
    console.error('\n💡 你可以手动下载二进制文件:');
    console.error(`   https://github.com/Xuzan9396/excel_mcp/releases\n`);
    process.exit(1);
  }
}

// 运行安装
install();
