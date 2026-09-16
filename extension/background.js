// Clicking the toolbar icon opens the hub as a resizable side panel docked to
// the current window (native browser behavior — drag its inner edge to
// resize). The app itself has a "maximize" button that opens the same page
// as a full tab when more room is wanted.
//
// Belt and suspenders: setPanelBehavior is the normal way to make the toolbar
// icon open the side panel automatically, but if that ever silently fails
// (extension update timing, browser quirk), the explicit onClicked handler
// below opens the panel directly too — so a click should never fall through
// to doing nothing, or opening a tab.

chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch((err) => {
    console.warn("setPanelBehavior failed, relying on onClicked fallback:", err);
  });

  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: "open-d365-discovery-hub",
      title: "Open D365 Discovery Hub",
      contexts: ["page", "action"]
    });
  });
});

chrome.action.onClicked.addListener((tab) => {
  chrome.sidePanel.open({ windowId: tab.windowId }).catch((err) => {
    console.error("Failed to open side panel:", err);
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if(info.menuItemId === "open-d365-discovery-hub" && tab){
    chrome.sidePanel.open({ windowId: tab.windowId }).catch((err) => {
      console.error("Failed to open side panel from context menu:", err);
    });
  }
});
