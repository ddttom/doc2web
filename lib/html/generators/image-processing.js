// File: lib/html/generators/image-processing.js
// Image processing and extraction utilities

const mammoth = require("mammoth");
const path = require("path");
const fs = require("fs");

/**
 * Create image options for mammoth conversion
 */
function createImageOptions() {
  return {
    convertImage: mammoth.images.imgElement(function (image) {
      return image.read("base64").then(function (imageBuffer) {
        try {
          const extension = image.contentType.split("/")[1] || "png";
          const altTextSanitized = (image.altText || `image-${Date.now()}`)
            .replace(/[^\w-]/g, "_")
            .substring(0, 50);
          const filename = `${altTextSanitized}.${extension}`;
          return {
            src: `./images/${filename}`,
            alt: image.altText || "Document image",
            className: "docx-image",
          };
        } catch (imgError) {
          console.error("Error processing single image:", imgError);
          return {
            src: `./images/error-image-${Date.now()}.png`,
            alt: "Error processing image",
            className: "docx-image docx-image-error",
          };
        }
      });
    }),
  };
}

/**
 * Extract images from DOCX file to output directory
 */
async function extractImagesFromDocx(docxPath, outputDir) {
  try {
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    const imageOptions = {
      path: docxPath,
      convertImage: mammoth.images.imgElement(function (image) {
        return image.read("base64").then(function (imageBuffer) {
          try {
            const extension = image.contentType.split("/")[1] || "png";
            const sanitizedAltText = (image.altText || `image-${Date.now()}`)
              .replace(/[^\w-]/g, "_")
              .substring(0, 50);
            const filename = `${sanitizedAltText}.${extension}`;
            const imagePath = path.join(outputDir, filename);
            fs.writeFileSync(imagePath, Buffer.from(imageBuffer, "base64"));
            return {
              src: `./images/${filename}`,
              alt: image.altText || "Document image",
            };
          } catch (imgError) {
            console.error("Error saving image:", imgError);
            return {
              src: `./images/error-img-${Date.now()}.${extension}`,
              alt: "Error saving image",
            };
          }
        });
      }),
    };
    await mammoth.convertToHtml(imageOptions);
    console.log(`Images extracted to ${outputDir}`);
  } catch (error) {
    console.error("Error extracting images:", error.message, error.stack);
  }
}

module.exports = {
  createImageOptions,
  extractImagesFromDocx,
};