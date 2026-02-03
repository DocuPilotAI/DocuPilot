import { useState } from 'react';
import { X } from 'lucide-react';

interface ChatTabProps {
  id: string;
  name: string;
  isActive: boolean;
  onSelect: () => void;
  onClose: () => void;
}

export function ChatTab({ name, isActive, onSelect, onClose }: ChatTabProps) {
  const [isHovered, setIsHovered] = useState(false);

  const handleClose = (e: React.MouseEvent) => {
    e.stopPropagation();
    onClose();
  };

  return (
    <button
      onClick={onSelect}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      className={`
        flex items-center gap-2 px-3 py-1.5 text-[12px] rounded-md
        transition-colors relative group
        ${isActive 
          ? 'bg-[#f5f5f5] text-[#0a0a0a]' 
          : 'text-[#737373] hover:bg-[#fafafa] hover:text-[#0a0a0a]'
        }
      `}
    >
      <span className="truncate max-w-[120px]">{name}</span>
      {isHovered && (
        <span
          role="button"
          tabIndex={0}
          onClick={handleClose}
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              handleClose(e as any);
            }
          }}
          className="flex items-center justify-center w-4 h-4 rounded hover:bg-[#e5e5e5] transition-colors cursor-pointer"
          aria-label="Close tab"
        >
          <X className="w-3 h-3" />
        </span>
      )}
    </button>
  );
}
