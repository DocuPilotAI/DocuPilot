// Office.js type declarations
declare namespace Office {
  interface OfficeInfo {
    host: HostType;
    platform: PlatformType;
  }

  enum HostType {
    Word = "Word",
    Excel = "Excel",
    PowerPoint = "PowerPoint",
    Outlook = "Outlook",
    OneNote = "OneNote",
    Project = "Project",
    Access = "Access",
  }

  enum PlatformType {
    PC = "PC",
    OfficeOnline = "OfficeOnline",
    Mac = "Mac",
    iOS = "iOS",
    Android = "Android",
    Universal = "Universal",
  }

  const context: {
    host: HostType;
    platform: PlatformType;
    requirements: {
      isSetSupported(name: string, version?: string): boolean;
    };
  };

  function onReady(callback: (info: OfficeInfo) => void): Promise<OfficeInfo>;
}

declare const Office: typeof Office;
