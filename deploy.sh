#!/bin/bash
# AW2026 Report 배포 스크립트
# 사용법: ./deploy.sh

set -e

echo "=== AW2026 Report 배포 시작 ==="

# 변경사항 확인
if [ -z "$(git status --porcelain)" ]; then
    echo "변경사항이 없습니다."
    exit 0
fi

# 변경된 파일 목록
echo "변경된 파일:"
git status --short

# 스테이징 및 커밋
git add -A
echo ""
read -p "커밋 메시지를 입력하세요: " msg
if [ -z "$msg" ]; then
    msg="Update report $(date +%Y-%m-%d_%H:%M)"
fi

git commit -m "$msg"
git push origin main

echo ""
echo "=== 배포 완료 ==="
echo "1~2분 후 확인: https://mojaemin.github.io/aw2026-report/"
