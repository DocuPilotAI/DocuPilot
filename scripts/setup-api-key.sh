#!/bin/bash

# DocuPilot API 密钥配置脚本
# 用于快速设置 Anthropic API 密钥

set -e

echo "🔑 DocuPilot API 密钥配置向导"
echo "================================"
echo ""

# 检查是否在正确的目录
if [ ! -f "package.json" ]; then
    echo "❌ 错误: 请在 DocuPilot 项目根目录运行此脚本"
    exit 1
fi

# 检查是否已有配置文件
if [ -f ".env.local" ]; then
    echo "⚠️  发现现有的 .env.local 文件"
    echo ""
    
    # 检查是否已配置 API 密钥
    if grep -q "^ANTHROPIC_API_KEY=sk-ant-" .env.local 2>/dev/null; then
        echo "✅ 已配置 API 密钥"
        
        # 读取现有配置
        current_key=$(grep "^ANTHROPIC_API_KEY=" .env.local | cut -d'=' -f2)
        masked_key="${current_key:0:20}...${current_key: -4}"
        echo "   当前密钥: $masked_key"
        echo ""
        
        read -p "是否要更新 API 密钥? (y/N): " update_key
        if [[ ! "$update_key" =~ ^[Yy]$ ]]; then
            echo "保持现有配置"
            exit 0
        fi
    else
        echo "⚠️  未找到有效的 API 密钥"
        echo ""
    fi
    
    # 备份现有文件
    backup_file=".env.local.backup.$(date +%Y%m%d-%H%M%S)"
    cp .env.local "$backup_file"
    echo "📋 已备份现有配置到: $backup_file"
    echo ""
fi

# 获取 API 密钥
echo "📝 请输入您的 Anthropic API 密钥"
echo "   (从 https://console.anthropic.com/ 获取)"
echo ""
read -p "API 密钥 (sk-ant-api03-...): " api_key

# 验证密钥格式
if [[ ! "$api_key" =~ ^sk-ant- ]]; then
    echo ""
    echo "❌ 错误: API 密钥格式不正确"
    echo "   密钥应以 'sk-ant-' 开头"
    exit 1
fi

# 询问可选配置
echo ""
echo "⚙️  可选配置（直接按回车跳过）"
echo ""

read -p "自定义 API 端点 (留空使用默认): " base_url
read -p "指定模型 (留空使用默认): " model

# 生成配置文件
echo ""
echo "📝 生成配置文件..."

cat > .env.local << EOF
# ===== 必需配置 =====
# Anthropic API 密钥
ANTHROPIC_API_KEY=$api_key

# ===== 可选配置 =====
EOF

# 添加可选配置
if [ -n "$base_url" ]; then
    echo "# Anthropic API 基础 URL" >> .env.local
    echo "ANTHROPIC_BASE_URL=$base_url" >> .env.local
else
    echo "# Anthropic API 基础 URL（可选）" >> .env.local
    echo "# ANTHROPIC_BASE_URL=https://api.anthropic.com" >> .env.local
fi

echo "" >> .env.local

if [ -n "$model" ]; then
    echo "# Claude 模型" >> .env.local
    echo "ANTHROPIC_MODEL=$model" >> .env.local
else
    echo "# Claude 模型（可选）" >> .env.local
    echo "# ANTHROPIC_MODEL=claude-3-5-sonnet-20241022" >> .env.local
fi

echo "" >> .env.local
echo "# 启用测试面板（开发环境）" >> .env.local
echo "NEXT_PUBLIC_ENABLE_TEST_PANEL=1" >> .env.local

echo ""
echo "✅ 配置文件已创建: .env.local"
echo ""

# 测试 API 密钥（可选）
echo "🧪 是否测试 API 密钥?"
read -p "运行测试? (y/N): " run_test

if [[ "$run_test" =~ ^[Yy]$ ]]; then
    echo ""
    echo "📡 测试 API 连接..."
    
    # 简单的 curl 测试
    response=$(curl -s -o /dev/null -w "%{http_code}" \
        https://api.anthropic.com/v1/messages \
        -H "x-api-key: $api_key" \
        -H "anthropic-version: 2023-06-01" \
        -H "content-type: application/json" \
        -d '{
            "model": "claude-3-5-sonnet-20241022",
            "max_tokens": 1,
            "messages": [{"role": "user", "content": "test"}]
        }' 2>/dev/null || echo "000")
    
    if [ "$response" = "200" ]; then
        echo "✅ API 密钥有效！"
    elif [ "$response" = "401" ]; then
        echo "❌ API 密钥无效或已过期"
        echo "   请在 https://console.anthropic.com/ 检查您的密钥"
    elif [ "$response" = "000" ]; then
        echo "⚠️  无法连接到 API（可能是网络问题）"
        echo "   但配置文件已创建，可以继续"
    else
        echo "⚠️  收到意外响应: HTTP $response"
        echo "   配置文件已创建，请手动验证"
    fi
fi

echo ""
echo "🚀 下一步:"
echo "   1. 确保配置正确: cat .env.local"
echo "   2. 启动开发服务器: npm run dev:https"
echo "   3. 在 Office 中打开插件并测试"
echo ""
echo "📚 更多帮助:"
echo "   - 配置文档: doc/API_KEY_SETUP.md"
echo "   - 快速开始: doc/QUICK_START.md"
echo "   - 故障排除: doc/TROUBLESHOOTING.md"
echo ""
echo "✨ 配置完成！"
