const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('ekp', {
  catalog: () => ipcRenderer.invoke('catalog'), state: () => ipcRenderer.invoke('state:get'), saveState: state => ipcRenderer.invoke('state:set', state),
  catalogMetadata: () => ipcRenderer.invoke('catalog:metadata'), updateHistory: () => ipcRenderer.invoke('updates:history'),
  databaseUpdateStatus: () => ipcRenderer.invoke('database:update-status'), checkDatabaseUpdate: () => ipcRenderer.invoke('database:check-update'), installDatabaseUpdate: manifest => ipcRenderer.invoke('database:install-update', manifest), onDatabaseUpdateProgress: callback => { const listener = (_event, value) => callback(value); ipcRenderer.on('database:update-progress', listener); return () => ipcRenderer.removeListener('database:update-progress', listener); },
  notes: () => ipcRenderer.invoke('notes:list'), saveNote: note => ipcRenderer.invoke('notes:save', note), deleteNote: id => ipcRenderer.invoke('notes:delete', id),
  publicNotes: () => ipcRenderer.invoke('public-notes:list'), publicNotesUpdateStatus: () => ipcRenderer.invoke('public-notes:update-status'), checkPublicNotesUpdate: () => ipcRenderer.invoke('public-notes:check-update'), installPublicNotesUpdate: () => ipcRenderer.invoke('public-notes:install-update'), onPublicNotesUpdateProgress: callback => { const listener = (_event, value) => callback(value); ipcRenderer.on('public-notes:update-progress', listener); return () => ipcRenderer.removeListener('public-notes:update-progress', listener); },
  adminLogin: credentials => ipcRenderer.invoke('admin:login', credentials), adminLogout: () => ipcRenderer.invoke('admin:logout'), exportNotesLibrary: payload => ipcRenderer.invoke('admin:export-library', payload),
  addAttachments: id => ipcRenderer.invoke('attachments:add', id), deleteAttachment: file => ipcRenderer.invoke('attachments:delete', file), readAttachment: file => ipcRenderer.invoke('attachments:read', file), openFile: file => ipcRenderer.invoke('file:open', file), openDbn: file => ipcRenderer.invoke('dbn:open', file), readPdf: file => ipcRenderer.invoke('dbn:read-pdf', file),
  openExternal: url => ipcRenderer.invoke('external:open', url),
  window: { minimize: () => ipcRenderer.invoke('window:minimize'), maximize: () => ipcRenderer.invoke('window:maximize'), close: () => ipcRenderer.invoke('window:close') }
});
