"use client";

import { useState, useEffect } from "react";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "./ui/dialog";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./ui/tabs";
import { Label } from "./ui/label";
import { Input } from "./ui/input";
import { Button } from "./ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./ui/select";
import { Bot, AlertCircle } from "lucide-react";

interface SettingsDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

interface ApiSettings {
  apiKey: string;
  apiUrl: string;
  modelName: string;
  language?: string;
}

const STORAGE_KEY = "docupilot:api-settings";

const DEFAULT_MODELS = [
  { value: "claude-haiku-4-5-20251001", label: "Claude 4.5 Haiku" },
  { value: "claude-haiku-4-5-20251001-thinking", label: "Claude 4.5 Haiku (Thinking)" },
  { value: "claude-sonnet-4-5-20250929", label: "Claude 4.5 Sonnet (推荐)" },
  { value: "claude-sonnet-4-5-20250929-thinking", label: "Claude 4.5 Sonnet (Thinking)" },
  { value: "claude-opus-4-5-20251101", label: "Claude 4.5 Opus" },
  { value: "claude-opus-4-5-20251101-thinking", label: "Claude 4.5 Opus (Thinking)" },
];

const LANGUAGE_OPTIONS = [
  { value: "default", label: "默认 (Default)" },
  { value: "chinese", label: "简体中文 (Chinese)" },
  { value: "english", label: "English" },
  { value: "japanese", label: "日本語 (Japanese)" },
  { value: "spanish", label: "Español (Spanish)" },
  { value: "french", label: "Français (French)" },
  { value: "german", label: "Deutsch (German)" },
  { value: "korean", label: "한국어 (Korean)" },
];

// 加载设置
export function loadApiSettings(): ApiSettings | null {
  if (typeof window === "undefined") return null;
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return null;
    return JSON.parse(stored);
  } catch (error) {
    console.error("Failed to load API settings:", error);
    return null;
  }
}

// 保存设置
export function saveApiSettings(settings: ApiSettings): void {
  if (typeof window === "undefined") return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch (error) {
    console.error("Failed to save API settings:", error);
  }
}

// 清除设置
export function clearApiSettings(): void {
  if (typeof window === "undefined") return;
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (error) {
    console.error("Failed to clear API settings:", error);
  }
}

export function SettingsDialog({ open, onOpenChange }: SettingsDialogProps) {
  const [apiKey, setApiKey] = useState("");
  const [apiUrl, setApiUrl] = useState("");
  const [modelName, setModelName] = useState("claude-sonnet-4-5-20250929");
  const [customModel, setCustomModel] = useState("");
  const [showCustomModel, setShowCustomModel] = useState(false);
  const [language, setLanguage] = useState("default");
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [showWarning, setShowWarning] = useState(true);

  // 加载设置
  useEffect(() => {
    if (open) {
      const settings = loadApiSettings();
      if (settings) {
        setApiKey(settings.apiKey || "");
        setApiUrl(settings.apiUrl || "");
        setModelName(settings.modelName || "claude-sonnet-4-5-20250929");
        setLanguage(settings.language || "default");
        
        // 检查是否是自定义模型
        const isDefaultModel = DEFAULT_MODELS.some(m => m.value === settings.modelName);
        if (!isDefaultModel && settings.modelName) {
          setCustomModel(settings.modelName);
          setShowCustomModel(true);
        }
      }
      
      // 从服务器加载语言设置
      fetch("/api/settings/language")
        .then((res) => res.json())
        .then((data) => {
          if (data.language) {
            setLanguage(data.language);
          }
        })
        .catch((error) => {
          console.error("Failed to load language setting:", error);
        });
      
      setSaveSuccess(false);
    }
  }, [open]);

  const handleSave = async () => {
    const finalModelName = showCustomModel && customModel.trim() 
      ? customModel.trim() 
      : modelName;

    const settings: ApiSettings = {
      apiKey: apiKey.trim(),
      apiUrl: apiUrl.trim(),
      modelName: finalModelName,
      language: language,
    };

    // 保存到 localStorage
    saveApiSettings(settings);
    
    // 保存语言设置到 .claude/settings.local.json
    // 只在选择了非默认语言时才保存
    if (language && language !== "default") {
      try {
        await fetch("/api/settings/language", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ language }),
        });
      } catch (error) {
        console.error("Failed to save language setting:", error);
      }
    }
    
    setSaveSuccess(true);
    
    setTimeout(() => {
      setSaveSuccess(false);
      onOpenChange(false);
    }, 1500);
  };

  const handleClear = () => {
    if (confirm("确定要清除所有 API 设置吗？这将使用环境变量中的配置。")) {
      clearApiSettings();
      setApiKey("");
      setApiUrl("");
      setModelName("claude-sonnet-4-5-20250929");
      setCustomModel("");
      setShowCustomModel(false);
      setLanguage("default");
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Bot className="w-5 h-5" />
            设置
          </DialogTitle>
          <DialogDescription>
            配置 Anthropic API 参数。留空则使用环境变量配置。
          </DialogDescription>
        </DialogHeader>

        <Tabs defaultValue="model" className="w-full">
          <TabsList className="grid w-full grid-cols-1">
            <TabsTrigger value="model">API 配置</TabsTrigger>
          </TabsList>

          <TabsContent value="model" className="space-y-4 mt-4">
            {/* 安全提示 */}
            {showWarning && (
              <div className="flex items-start gap-2 p-3 bg-amber-50 border border-amber-200 rounded-md">
                <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 flex-shrink-0" />
                <div className="flex-1 text-xs text-amber-800">
                  <p className="font-medium mb-1">安全提示</p>
                  <p>API Key 将保存在浏览器本地存储中。请勿在公共设备上保存敏感信息。</p>
                  <button
                    onClick={() => setShowWarning(false)}
                    className="text-amber-600 hover:text-amber-700 underline mt-1"
                  >
                    知道了
                  </button>
                </div>
              </div>
            )}

            {/* API Key */}
            <div className="space-y-2">
              <Label htmlFor="api-key">API Key</Label>
              <Input
                id="api-key"
                type="password"
                placeholder="sk-ant-..."
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
              />
              <p className="text-xs text-gray-500">
                留空使用环境变量 ANTHROPIC_API_KEY
              </p>
            </div>

            {/* API URL */}
            <div className="space-y-2">
              <Label htmlFor="api-url">API 端点 (可选)</Label>
              <Input
                id="api-url"
                type="url"
                placeholder="https://api.anthropic.com"
                value={apiUrl}
                onChange={(e) => setApiUrl(e.target.value)}
              />
              <p className="text-xs text-gray-500">
                留空使用 Anthropic 官方地址或环境变量 ANTHROPIC_BASE_URL
              </p>
            </div>

            {/* Model Name */}
            <div className="space-y-2">
              <Label htmlFor="model-name">模型</Label>
              {!showCustomModel ? (
                <Select value={modelName} onValueChange={setModelName}>
                  <SelectTrigger id="model-name">
                    <SelectValue placeholder="选择模型" />
                  </SelectTrigger>
                  <SelectContent>
                    {DEFAULT_MODELS.map((model) => (
                      <SelectItem key={model.value} value={model.value}>
                        {model.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              ) : (
                <Input
                  id="custom-model"
                  type="text"
                  placeholder="claude-sonnet-4-5-20250929"
                  value={customModel}
                  onChange={(e) => setCustomModel(e.target.value)}
                />
              )}
              <button
                onClick={() => setShowCustomModel(!showCustomModel)}
                className="text-xs text-blue-600 hover:text-blue-700 underline"
              >
                {showCustomModel ? "使用预设模型" : "使用自定义模型"}
              </button>
            </div>

            {/* Language */}
            <div className="space-y-2">
              <Label htmlFor="language">回复语言</Label>
              <Select value={language} onValueChange={setLanguage}>
                <SelectTrigger id="language">
                  <SelectValue placeholder="选择语言" />
                </SelectTrigger>
                <SelectContent>
                  {LANGUAGE_OPTIONS.map((lang) => (
                    <SelectItem key={lang.value} value={lang.value}>
                      {lang.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <p className="text-xs text-gray-500">
                Claude 将使用您选择的语言进行回复
              </p>
            </div>

            {/* 优先级说明 */}
            <div className="p-3 bg-gray-50 border border-gray-200 rounded-md">
              <p className="text-xs text-gray-600 font-medium mb-1">配置优先级</p>
              <ol className="text-xs text-gray-600 space-y-1 list-decimal list-inside">
                <li>前端设置（当前页面配置）</li>
                <li>.env.local 文件配置</li>
                <li>系统环境变量</li>
              </ol>
            </div>

            {/* 保存成功提示 */}
            {saveSuccess && (
              <div className="p-3 bg-green-50 border border-green-200 rounded-md text-sm text-green-800">
                ✓ 设置已保存
              </div>
            )}

            {/* 按钮 */}
            <div className="flex justify-between pt-4">
              <Button
                variant="outline"
                onClick={handleClear}
                className="text-red-600 hover:text-red-700 hover:bg-red-50"
              >
                清除设置
              </Button>
              <div className="flex gap-2">
                <Button
                  variant="outline"
                  onClick={() => onOpenChange(false)}
                >
                  取消
                </Button>
                <Button onClick={handleSave}>
                  保存
                </Button>
              </div>
            </div>
          </TabsContent>
        </Tabs>
      </DialogContent>
    </Dialog>
  );
}
