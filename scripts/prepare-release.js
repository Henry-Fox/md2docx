const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// 首先运行构建命令
console.log('Building project...');
execSync('npm run build', { stdio: 'inherit' });

// 创建发布目录
const releaseDir = path.join(__dirname, '../release');
if (!fs.existsSync(releaseDir)) {
    fs.mkdirSync(releaseDir);
}

// 复制 dist 目录中的所有文件到 release 目录
const distDir = path.join(__dirname, '../dist');
if (fs.existsSync(distDir)) {
    execSync(`xcopy "${distDir}" "${releaseDir}" /E /I /Y`);
    console.log('Copied: dist files to release directory');
}

// 复制其他必要文件
const filesToCopy = [
    'LICENSE',
    'README.md'
];

filesToCopy.forEach(file => {
    const source = path.join(__dirname, '..', file);
    const target = path.join(releaseDir, file);

    if (fs.existsSync(source)) {
        fs.copyFileSync(source, target);
        console.log(`Copied: ${file}`);
    } else {
        console.warn(`Warning: ${file} not found`);
    }
});

// 创建版本信息文件
const version = require('../package.json').version;
const versionInfo = {
    version,
    buildDate: new Date().toISOString(),
    commit: execSync('git rev-parse HEAD').toString().trim()
};

fs.writeFileSync(
    path.join(releaseDir, 'version.json'),
    JSON.stringify(versionInfo, null, 2)
);

// 创建使用说明文件
const usageGuide = `# 本地使用说明

## 使用方法
1. 解压下载的文件
2. 直接用浏览器打开 index.html 文件即可使用

## 在线使用
您也可以直接访问在线版本：https://henry-fox.github.io/md2docx/
`;

fs.writeFileSync(
    path.join(releaseDir, 'USAGE.md'),
    usageGuide
);

// 创建 zip 文件
console.log('Creating zip file...');
const zipFile = path.join(__dirname, '..', `md2docx-v${version}.zip`);
if (fs.existsSync(zipFile)) {
    fs.unlinkSync(zipFile);
}

// 使用 PowerShell 创建 zip 文件
execSync(`powershell Compress-Archive -Path "${releaseDir}\\*" -DestinationPath "${zipFile}" -Force`);

console.log('Release package prepared successfully!');
console.log(`Zip file created at: ${zipFile}`);
