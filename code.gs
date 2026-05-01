function doPost(e) {
	const payload = JSON.parse(e.postData.contents);

	const rootFolderId = payload.rootFolderId;
	const folderPath = payload.folderPath || [];
	const filename = payload.filename;
	const mime = payload.mime;
	const base64 = payload.base64;

	let folder = DriveApp.getFolderById(rootFolderId);

	folderPath.forEach(name => {
		folder = getOrCreateFolder(folder, name);
	});

	const bytes = Utilities.base64Decode(base64);
	const blob = Utilities.newBlob(bytes, mime, filename);

	const existing = folder.getFilesByName(filename);
	while (existing.hasNext()) {
		existing.next().setTrashed(true);
	}

	const file = folder.createFile(blob);

	return ContentService
		.createTextOutput(JSON.stringify({
			status: "ok",
			url: file.getUrl(),
			id: file.getId()
		}))
		.setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder(parent, name) {
	const folders = parent.getFoldersByName(name);
	if (folders.hasNext()) {
		return folders.next();
	}
	return parent.createFolder(name);
}
