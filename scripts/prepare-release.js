const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// 创建发布目录
const releaseDir = path.join(__dirname, '../release');
if (!fs.existsSync(releaseDir)) {
    fs.mkdirSync(releaseDir);
}

// 需要复制的文件和目录
const filesToCopy = [
    'index.html',
    'css',
    'js',
    'img',
    'favicon.ico',
    'icon.svg',
    'icon.png',
    'site.webmanifest',
    'LICENSE',
    'README.md'
];

// 复制文件
filesToCopy.forEach(file => {
    const source = path.join(__dirname, '..', file);
    const target = path.join(releaseDir, file);

    if (fs.existsSync(source)) {
        if (fs.lstatSync(source).isDirectory()) {
            // 复制目录
            execSync(`xcopy "${source}" "${target}" /E /I /Y`);
        } else {
            // 复制文件
            fs.copyFileSync(source, target);
        }
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

console.log('Release package prepared successfully!');
