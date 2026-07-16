# Publishing checklist

Local repo is ready (`master` @ `8fafc14`). Auth was not available in this environment
(`npm whoami` → ENEEDAUTH, GitHub SSH → Permission denied).

## 1. GitHub

```bash
cd c:\Projects\ai-document-parser
# Create empty public repo "ai-document-parser" under novitsky-413x, then:
git branch -M main
git remote add origin git@github.com:novitsky-413x/ai-document-parser.git
git push -u origin main
```

Or with HTTPS + personal access token:

```bash
git remote add origin https://github.com/novitsky-413x/ai-document-parser.git
git push -u origin main
```

## 2. npm

```bash
cd c:\Projects\ai-document-parser
npm login
npm run prepublishOnly   # licenses + build + tests
npm publish --access public
```

After publish, point WATCHER at the registry version:

```json
"ai-document-parser": "^0.1.0"
```

instead of `"file:../ai-document-parser"`.
