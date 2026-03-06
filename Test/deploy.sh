#!/bin/bash

echo ""
echo "🚀 GitHub Pages 배포를 시작합니다..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

CHANGED=$(git status --porcelain)
if [ -z "$CHANGED" ]; then
  echo "⚠️  변경된 파일이 없습니다. 수정 후 다시 실행해주세요."
  exit 0
fi

echo "📄 변경된 파일:"
git status --short
echo ""

TIMESTAMP=$(date '+%Y-%m-%d %H:%M')
git add .
git commit -m "시안 업데이트 - $TIMESTAMP"
git push origin main

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "✅ 배포 완료!"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""