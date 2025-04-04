/**
 * ResizeObserverをグローバルに定義する型宣言
 */
interface Window {
  ResizeObserver: typeof ResizeObserver;
}
