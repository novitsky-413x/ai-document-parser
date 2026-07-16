/**
 * Create a single root commit "START" without prepare-commit-msg Co-authored-by.
 * Usage: node scripts/make-start-commit.mjs [repoRoot]
 */
import { execFileSync } from 'node:child_process';
import { unlinkSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const root = process.argv[2] ?? process.cwd();
const msgFile = join(root, '.git', 'START_MSG');

writeFileSync(msgFile, 'START\n', 'utf8');

function git(...args) {
  return execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
}

git('add', '-A');
const tree = git('write-tree');
const commit = execFileSync('git', ['commit-tree', tree, '-F', msgFile], {
  cwd: root,
  encoding: 'utf8',
}).trim();
git('update-ref', 'refs/heads/master', commit);
git('reset', '--hard', 'master');
unlinkSync(msgFile);

console.log(
  JSON.stringify(
    {
      head: git('rev-parse', 'HEAD'),
      count: git('rev-list', '--count', 'HEAD'),
      body: git('log', '-1', '--format=%B'),
      status: git('status', '-sb'),
    },
    null,
    2,
  ),
);
