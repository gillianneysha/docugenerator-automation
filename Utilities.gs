/**
 * Utilities.gs
 * Small shared helpers that don't belong to a specific service.
 */

function sanitizeFileName_(name) {
  return name.replace(/[\\/:*?"<>|]/g, '-').trim();
}
