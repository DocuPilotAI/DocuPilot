import { motion, AnimatePresence } from 'framer-motion';
import { Clock, MessageSquare, Trash2 } from 'lucide-react';

interface HistoryItem {
  id: string;
  name: string;
  timestamp: string;
  preview: string;
}

interface HistoryPanelProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectHistory: (id: string) => void;
  onDeleteHistory: (id: string) => void;
  historyItems: HistoryItem[];
}

export function HistoryPanel({ 
  isOpen, 
  onClose, 
  onSelectHistory, 
  onDeleteHistory,
  historyItems 
}: HistoryPanelProps) {
  return (
    <AnimatePresence>
      {isOpen && (
        <>
          {/* Backdrop */}
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="fixed inset-0 bg-black/20 z-40"
          />
          
          {/* Panel */}
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: 20 }}
            transition={{ duration: 0.15 }}
            className="fixed right-0 top-0 bottom-0 w-80 bg-[#ffffff] border-l border-[rgba(0,0,0,0.08)] shadow-xl z-50 flex flex-col"
          >
            {/* Header */}
            <div className="px-4 py-3 border-b border-[rgba(0,0,0,0.08)]">
              <div className="flex items-center gap-2">
                <Clock className="w-4 h-4 text-[#737373]" />
                <h2 className="text-[13px] font-medium text-[#0a0a0a]">History</h2>
              </div>
            </div>

            {/* History List */}
            <div className="flex-1 overflow-y-auto p-2">
              {historyItems.map((item) => (
                <div
                  key={item.id}
                  className="group w-full p-3 rounded-lg hover:bg-[#f5f5f5] transition-colors mb-1 relative"
                >
                  <button
                    onClick={() => {
                      onSelectHistory(item.id);
                      onClose();
                    }}
                    className="w-full text-left"
                  >
                    <div className="flex items-start gap-2">
                      <MessageSquare className="w-4 h-4 text-[#737373] mt-0.5 flex-shrink-0" />
                      <div className="flex-1 min-w-0 pr-6">
                        <h3 className="text-[12px] font-medium text-[#0a0a0a] truncate">
                          {item.name}
                        </h3>
                        <p className="text-[11px] text-[#737373] truncate mt-0.5">
                          {item.preview}
                        </p>
                        <p className="text-[10px] text-[#a3a3a3] mt-1">
                          {item.timestamp}
                        </p>
                      </div>
                    </div>
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDeleteHistory(item.id);
                    }}
                    className="absolute right-3 top-3 opacity-0 group-hover:opacity-100 transition-opacity p-1 hover:bg-[#e5e5e5] rounded"
                    aria-label="删除对话"
                    title="删除对话"
                  >
                    <Trash2 className="w-3.5 h-3.5 text-[#737373]" />
                  </button>
                </div>
              ))}
            </div>
          </motion.div>
        </>
      )}
    </AnimatePresence>
  );
}
