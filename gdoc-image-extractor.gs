/*
 * ============================================================
 * GOOGLE DOC IMAGE EXTRACTOR
 * ============================================================
 *
 * - Extracts EVERY image occurrence from the document.
 * - Repeated/duplicate images in the document are preserved.
 * - Rerunning does NOT extract the same occurrence again.
 * - Images are saved beside the Google Doc.
 * - No explicit PNG/JPEG conversion is performed.
 * - Existing files are never overwritten.
 * ============================================================
 */


/* ------------------------------------------------------------
 * UI
 * ------------------------------------------------------------ */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Image Extractor')
    .addItem('Extract Images', 'uiExtractImages')
    .addSeparator()
    .addItem('Help', 'uiImageExtractorHelp')
    .addToUi();
}


function uiExtractImages() {
  const ui = DocumentApp.getUi();

  try {
    const doc = DocumentApp.getActiveDocument();

    if (!doc) {
      throw new Error('No active Google Doc found.');
    }

    const result = getDocImages(doc.getId());

    let message =
      `Images in document: ${result.imageCount}\n` +
      `Newly extracted: ${result.createdCount}\n` +
      `Already extracted: ${result.skippedCount}`;

    if (result.created.length) {
      message += '\n\nCreated:\n';

      result.created.forEach(file => {
        message += `• ${file.name}\n`;
      });
    }

    ui.alert(
      'Image Extraction Complete',
      message,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert(
      'Image Extraction Failed',
      err.message,
      ui.ButtonSet.OK
    );
  }
}


function uiImageExtractorHelp() {
  DocumentApp.getUi().alert(
    'Image Extractor',
    [
      'Choose:',
      '',
      'Image Extractor → Extract Images',
      '',
      'Every image occurrence in the document is extracted.',
      '',
      'If the same image appears several times in the document,',
      'each occurrence is extracted separately.',
      '',
      'Rerunning the extractor does not create another copy of',
      'an occurrence that was already extracted.',
      '',
      'Images are saved in the same Drive folder as this document.'
    ].join('\n'),
    DocumentApp.getUi().ButtonSet.OK
  );
}


/* ------------------------------------------------------------
 * MAIN
 * ------------------------------------------------------------ */

function getDocImages(sourceId) {
  if (!sourceId || typeof sourceId !== 'string') {
    throw new Error('A valid sourceId is required.');
  }

  /*
   * Prevent simultaneous runs from creating duplicate files.
   */
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sourceFile = DriveApp.getFileById(sourceId);

    const sourceName = sourceFile
      .getName()
      .replace(/\.[^.]+$/, '');

    const parents = sourceFile.getParents();

    if (!parents.hasNext()) {
      throw new Error(
        `Document "${sourceFile.getName()}" has no accessible parent folder.`
      );
    }

    /*
     * Save directly beside the document.
     */
    const folder = parents.next();
    const folderId = folder.getId();

    const doc = DocumentApp.openById(sourceId);
    const images = doc.getBody().getImages();

    const created = [];
    const skipped = [];

    /*
     * Tracks how many times each identical image has appeared
     * during THIS document traversal.
     *
     * Example:
     *
     * A
     * B
     * A
     *
     * becomes:
     *
     * A occurrence #1
     * B occurrence #1
     * A occurrence #2
     *
     * This is what allows repeated images to be preserved.
     */
    const hashOccurrences = {};

    images.forEach((image, index) => {
      const blob = image.getBlob();

      const mimeType = blob.getContentType();

      const extension =
        getImageExtension_(mimeType);

      const hash =
        getBlobHash_(blob);

      /*
       * Count this particular occurrence of this image.
       */
      hashOccurrences[hash] =
        (hashOccurrences[hash] || 0) + 1;

      const occurrenceNumber =
        hashOccurrences[hash];

      /*
       * Important:
       *
       * Deduplication is based on:
       *
       *   document
       *   + image content
       *   + occurrence number
       *
       * NOT image content alone.
       */
      const occurrenceKey =
        getOccurrenceKey_(
          sourceId,
          hash,
          occurrenceNumber
        );

      /*
       * Check whether THIS occurrence was already extracted.
       */
      const existing =
        getExistingOccurrence_(
          occurrenceKey,
          folderId,
          hash
        );

      if (existing) {
        skipped.push({
          id: existing.getId(),
          name: existing.getName(),
          sha256: hash,
          occurrence: occurrenceNumber
        });

        return;
      }

      /*
       * Migration / recovery convenience:
       *
       * If a file with the expected name already exists from an
       * older version of this script, check its actual content.
       *
       * This avoids unnecessarily recreating it.
       */
      const expectedFilename =
        `${sourceName}_${index + 1}.${extension}`;

      const existingNamedFile =
        findMatchingNamedFile_(
          folder,
          expectedFilename,
          hash
        );

      if (existingNamedFile) {
        saveOccurrence_(
          occurrenceKey,
          existingNamedFile.getId()
        );

        skipped.push({
          id: existingNamedFile.getId(),
          name: existingNamedFile.getName(),
          sha256: hash,
          occurrence: occurrenceNumber
        });

        return;
      }

      /*
       * Use the document position for readable filenames.
       */
      const baseName =
        `${sourceName}_${index + 1}`;

      /*
       * Never overwrite an unrelated existing file.
       */
      const filename =
        getUniqueFilename_(
          folder,
          baseName,
          extension
        );

      /*
       * Keep the blob as Google Docs exposes it.
       *
       * No getAs('image/png') conversion.
       */
      blob.setName(filename);

      const file =
        folder.createFile(blob);

      /*
       * Remember this particular occurrence.
       */
      saveOccurrence_(
        occurrenceKey,
        file.getId()
      );

      created.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: mimeType,
        size: file.getSize(),
        sha256: hash,
        occurrence: occurrenceNumber,
        url: file.getUrl()
      });
    });

    return {
      sourceId: sourceId,
      sourceName: sourceFile.getName(),

      destinationFolderId:
        folderId,

      imageCount:
        images.length,

      createdCount:
        created.length,

      skippedCount:
        skipped.length,

      created:
        created,

      skipped:
        skipped
    };

  } catch (err) {
    throw new Error(
      `Failed to extract images: ${err.message}`
    );

  } finally {
    lock.releaseLock();
  }
}


/* ------------------------------------------------------------
 * OCCURRENCE TRACKING
 * ------------------------------------------------------------ */

/*
 * Each repeated instance of an image gets its own key.
 *
 * Example:
 *
 * same SHA-256, first occurrence:
 *     ...:1
 *
 * same SHA-256, second occurrence:
 *     ...:2
 */
function getOccurrenceKey_(
  sourceId,
  hash,
  occurrence
) {
  return (
    `DOCIMAGE:${sourceId}:` +
    `${hash}:${occurrence}`
  );
}


function saveOccurrence_(key, fileId) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      key,
      fileId
    );
}


/*
 * Check whether an occurrence we previously extracted
 * still exists and still contains the expected bytes.
 */
function getExistingOccurrence_(
  key,
  folderId,
  expectedHash
) {
  const properties =
    PropertiesService.getScriptProperties();

  const fileId =
    properties.getProperty(key);

  if (!fileId) {
    return null;
  }

  try {
    const file =
      DriveApp.getFileById(fileId);

    /*
     * File was moved elsewhere.
     */
    if (!fileIsInFolder_(file, folderId)) {
      properties.deleteProperty(key);
      return null;
    }

    /*
     * Verify actual bytes before trusting the stored mapping.
     */
    const currentHash =
      getBlobHash_(file.getBlob());

    if (currentHash !== expectedHash) {
      properties.deleteProperty(key);
      return null;
    }

    return file;

  } catch (e) {
    /*
     * File was deleted or is inaccessible.
     */
    properties.deleteProperty(key);
    return null;
  }
}


/* ------------------------------------------------------------
 * OLD-FILE / MIGRATION CHECK
 * ------------------------------------------------------------ */

/*
 * Only checks files having the exact expected filename.
 *
 * It does NOT scan every file in the folder.
 */
function findMatchingNamedFile_(
  folder,
  filename,
  expectedHash
) {
  const files =
    folder.getFilesByName(filename);

  while (files.hasNext()) {
    const file =
      files.next();

    try {
      const mimeType =
        file.getMimeType();

      if (
        !mimeType ||
        !mimeType.startsWith('image/')
      ) {
        continue;
      }

      const hash =
        getBlobHash_(file.getBlob());

      if (hash === expectedHash) {
        return file;
      }

    } catch (e) {
      /*
       * Ignore unreadable candidates.
       */
    }
  }

  return null;
}


/* ------------------------------------------------------------
 * HELPERS
 * ------------------------------------------------------------ */

function fileIsInFolder_(file, folderId) {
  const parents =
    file.getParents();

  while (parents.hasNext()) {
    if (
      parents.next().getId() ===
      folderId
    ) {
      return true;
    }
  }

  return false;
}


function getBlobHash_(blob) {
  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      blob.getBytes()
    );

  return digest
    .map(byte => {
      const value =
        byte < 0
          ? byte + 256
          : byte;

      return value
        .toString(16)
        .padStart(2, '0');
    })
    .join('');
}


/*
 * Never overwrite a file.
 */
function getUniqueFilename_(
  folder,
  baseName,
  extension
) {
  let filename =
    `${baseName}.${extension}`;

  if (
    !folder
      .getFilesByName(filename)
      .hasNext()
  ) {
    return filename;
  }

  let suffix = 2;

  while (
    folder
      .getFilesByName(
        `${baseName}_${suffix}.${extension}`
      )
      .hasNext()
  ) {
    suffix++;
  }

  return (
    `${baseName}_${suffix}.${extension}`
  );
}


/*
 * Preserve the MIME type exposed by Google Docs and use
 * the corresponding extension.
 */
function getImageExtension_(mimeType) {
  const extensions = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/gif': 'gif',
    'image/bmp': 'bmp',
    'image/webp': 'webp',
    'image/tiff': 'tiff',
    'image/svg+xml': 'svg',
    'image/heic': 'heic',
    'image/heif': 'heif',
    'image/avif': 'avif'
  };

  if (extensions[mimeType]) {
    return extensions[mimeType];
  }

  if (
    mimeType &&
    mimeType.startsWith('image/')
  ) {
    const extension =
      mimeType
        .substring(6)
        .split('+')[0]
        .replace(
          /[^a-zA-Z0-9]/g,
          ''
        );

    if (extension) {
      return extension;
    }
  }

  throw new Error(
    `Unsupported image MIME type: ${mimeType}`
  );
}
