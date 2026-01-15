const fs = require('fs');
const path = require('path');

console.log('检查 nunjucks 依赖完整性...\n');

const dependencies = {
  'nunjucks': ['a-sync-waterfall', 'chokidar', 'commander', 'asap'],
  'xlsx': ['crc-32', 'adler-32', 'codepage', 'ssf']
};

function checkModule(moduleName) {
  const modulePaths = [
    path.join(__dirname, 'node_modules', moduleName),
    path.join(__dirname, 'node_modules', 'nunjucks', 'node_modules', moduleName)
  ];
  
  for (const modulePath of modulePaths) {
    if (fs.existsSync(modulePath)) {
      const pkgPath = path.join(modulePath, 'package.json');
      if (fs.existsSync(pkgPath)) {
        const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
        return {
          found: true,
          path: modulePath,
          version: pkg.version
        };
      }
      return { found: true, path: modulePath, version: 'unknown' };
    }
  }
  return { found: false };
}

// 检查所有依赖
Object.keys(dependencies).forEach(mainModule => {
  console.log(`检查 ${mainModule} 及其依赖：`);
  
  const mainResult = checkModule(mainModule);
  if (mainResult.found) {
    console.log(`  ✓ ${mainModule} (${mainResult.version})`);
    
    dependencies[mainModule].forEach(dep => {
      const depResult = checkModule(dep);
      if (depResult.found) {
        console.log(`    ✓ ${dep} (${depResult.version})`);
      } else {
        console.log(`    ✗ ${dep} 缺失，请运行: npm install ${dep} --save`);
      }
    });
  } else {
    console.log(`  ✗ ${mainModule} 缺失`);
  }
  console.log();
});