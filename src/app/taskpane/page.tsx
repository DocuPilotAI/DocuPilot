"use client";

import { useEffect, useState } from "react";
import { getOfficeHost, OfficeHostType } from "@/lib/office/host-detector";
import { getBridge, OfficeBridge } from "@/lib/office/bridge-factory";
import { Sidebar } from "@/components/Sidebar";
import { isInOfficeEnvironment, waitForOfficeReady } from "@/lib/office-loader";

export default function TaskpanePage() {
  const [hostType, setHostType] = useState<OfficeHostType>("unknown");
  const [bridge, setBridge] = useState<OfficeBridge | null>(null);
  const [isReady, setIsReady] = useState(false);

  useEffect(() => {
    let mounted = true;

    const initOffice = async () => {
      // 检测是否在 Office 环境中
      const inOffice = isInOfficeEnvironment();
      
      if (inOffice) {
        console.log("[TaskPane] Detected Office environment, waiting for Office.js...");
        
        try {
          const hostInfo = await waitForOfficeReady();
          
          if (!mounted) return;
          
          if (hostInfo) {
            console.log("[TaskPane] Office ready, host:", hostInfo);
            
            // 检测宿主类型
            const detectedHost = getOfficeHost();
            setHostType(detectedHost);
            
            // 获取对应的桥接实例
            const officeBridge = getBridge();
            setBridge(officeBridge);
            
            console.log("[TaskPane] Host type:", detectedHost);
            console.log("[TaskPane] Bridge ready:", officeBridge.isHostAvailable());
          } else {
            console.log("[TaskPane] Office not available, using mock mode");
            setHostType("excel"); // 默认使用 Excel 模式
          }
        } catch (error) {
          console.error("[TaskPane] Failed to initialize Office:", error);
          setHostType("excel"); // 出错时使用默认模式
        }
      } else {
        // 开发模式：非 Office 环境
        console.log("[TaskPane] Running outside Office, using development mode");
        setHostType("excel"); // 默认使用 Excel 模式进行开发
      }
      
      if (mounted) {
        setIsReady(true);
      }
    };

    initOffice();

    return () => {
      mounted = false;
    };
  }, []);

  if (!isReady) {
    return (
      <div className="flex items-center justify-center h-screen bg-gray-50">
        <div className="text-center">
          <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
          <p className="text-gray-600">正在初始化...</p>
        </div>
      </div>
    );
  }

  return (
    <main className="h-screen">
      <Sidebar 
        hostType={hostType} 
        bridge={bridge}
      />
    </main>
  );
}
