import * as Word from "office-js";

export interface ImageMetadata {
  type: number;
  url: string;
  siteid: string;
  ListItemid: string;
}

export class ImageHelper {
  /**
   * Embeds metadata into the alpha channel of an image file and returns a Base64 string of the modified image.
   * @param imageFile The original image file.
   * @param metadata The metadata to embed.
   * @returns Promise<string> - The Base64 string of the modified image.
   */
  private static async embedMetadataInImage(
    imageFile: File,
    metadata: ImageMetadata
  ): Promise<string> {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d")!;
    const img = new Image();

    const metadataString = JSON.stringify(metadata);
    const binaryData = Array.from(metadataString)
      .map((char) => char.charCodeAt(0).toString(2).padStart(8, "0"))
      .join("");

    return new Promise((resolve, reject) => {
      img.onload = () => {
        canvas.width = img.width;
        canvas.height = img.height;
        ctx.drawImage(img, 0, 0);

        const imageData = ctx.getImageData(0, 0, img.width, img.height);
        const data = imageData.data;

        // Embed binary data into the alpha channel
        for (let i = 0; i < binaryData.length && i * 4 + 3 < data.length; i++) {
          data[i * 4 + 3] = (data[i * 4 + 3] & ~1) | parseInt(binaryData[i], 2); // Modify alpha channel
        }

        ctx.putImageData(imageData, 0, 0);
        resolve(canvas.toDataURL());
      };

      img.onerror = (error) => reject(error);
      img.src = URL.createObjectURL(imageFile);
    });
  }

  /**
   * Inserts an image with embedded metadata into a Word document.
   * @param imageFile The image file to insert.
   * @param metadata The metadata to embed into the image.
   * @returns Promise<void>
   */
  public static async insertImageWithMetadata(
    imageFile: File,
    metadata: ImageMetadata
  ): Promise<void> {
    // Embed metadata into the image
    const modifiedImageBase64 = await this.embedMetadataInImage(imageFile, metadata);

    // Insert the modified image into the Word document
    await Word.run(async (context: Word.RequestContext) => {
      const docBody = context.document.body;
      docBody.insertInlinePictureFromBase64(modifiedImageBase64, Word.InsertLocation.end);
      await context.sync();
    });
  }
}
