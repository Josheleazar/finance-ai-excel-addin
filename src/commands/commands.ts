/*
 * Ribbon command function file.
 *
 * The stock template called Office.context.mailbox.item.notificationMessages which is
 * Outlook-only and throws in Excel. We register a safe no-op that simply completes the
 * event so the ribbon remains wired up. Add real command logic here when needed.
 */

/* global Office */

Office.onReady(() => {
  // No-op: the commands HTML is loaded on-demand when a ribbon button fires.
});

function action(event: Office.AddinCommands.Event): void {
  // TODO: implement a real ribbon action (e.g., open the task pane or refresh data).
  event.completed();
}

Office.actions.associate("action", action);
