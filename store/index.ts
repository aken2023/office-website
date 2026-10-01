import { create } from "zustand";
import { persist } from "zustand/middleware";
import { useSyncExternalStore } from "react";
import { EditorServer } from "@/utils/editor/server";
import {
  Language,
  Locale,
  LocaleExtend,
  standardizeLocale,
} from "@ziziyi/utils";
import { type OfficeTheme, type PluginMode } from "@/utils/editor/types";

/**
 * Resolves the language setting to an actual locale code.
 * If the language is set to "auto", it detects the browser's preferred language.
 */
function resolveLanguage(language: Language): Locale {
  if (language === LocaleExtend.Auto) {
    const browserLang =
      typeof navigator !== "undefined"
        ? navigator.language || (navigator as any).userLanguage
        : "en";
    return standardizeLocale(browserLang || "en");
  }
  return language as Locale;
}

/** 与 page.tsx 桥接约定一致 */
function notifyParentSaveFromStore(ok: boolean, error?: string) {
  if (typeof window === "undefined") return;
  const w = window as unknown as {
    __pendingParentSaveReply?: (ok: boolean, error?: string) => void;
    __pendingParentSaveId?: string;
  };
  if (w.__pendingParentSaveReply) {
    w.__pendingParentSaveReply(ok, error);
    w.__pendingParentSaveReply = undefined;
    w.__pendingParentSaveId = undefined;
  }
}

interface AppState {
  // Document State
  server: EditorServer;

  // Settings State
  language: Language;
  theme: OfficeTheme;
  plugins: PluginMode;

  // Actions
  setState: (
    state: Partial<Pick<AppState, "language" | "theme" | "plugins">>,
  ) => void;
}

export const useAppStore = create<AppState>()(
  persist(
    (set, get) => ({
      // Document Initial State
      server: new EditorServer({
        getState: () => get(),
        onRemoteSave: ({ ok, path, error }) => {
          // PUT 完成后通知坚果云父页面（关闭编辑等待用）
          notifyParentSaveFromStore(ok, error);
          if (ok) {
            console.log("[WebDAV] 已保存", path);
          } else {
            console.error("[WebDAV] 保存失败", error);
          }
        },
      }),

      // Settings Initial State
      language: LocaleExtend.Auto,
      theme: "theme-white",
      plugins: "featured",

      // Settings Actions
      setState: (newState) => set((state) => ({ ...state, ...newState })),
    }),
    {
      name: "office-state",
      // Only persist settings, skip server instance
      partialize: (state) => ({
        language: state.language,
        theme: state.theme,
        plugins: state.plugins,
      }),
    },
  ),
);

/**
 * Hook to check if persist rehydration has completed.
 * Returns false during SSR and before localStorage state is loaded,
 * then true once the persisted state has been applied.
 */
export function useHasHydrated(): boolean {
  return useSyncExternalStore(
    (callback) => {
      const unsub = useAppStore.persist.onFinishHydration(callback);
      return unsub;
    },
    () => useAppStore.persist.hasHydrated(),
    () => false, // SSR: always false
  );
}

/**
 * Hook to get the resolved language (reactive).
 * When language setting is "auto", returns the detected browser language.
 * Re-renders automatically when language setting changes.
 */
export function useResolvedLanguage(): Locale {
  return useAppStore((state) => resolveLanguage(state.language));
}