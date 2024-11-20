import * as Word from "office-js";

async function insertImageWithJson(imageFile, jsonObject) {
  // Embed JSON in the image
  const embedJsonInImage = async (imageFile, jsonObject) => {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    const img = new Image();

    const jsonString = JSON.stringify(jsonObject);
    const binaryData = Array.from(jsonString)
      .map((char) => char.charCodeAt(0).toString(2).padStart(8, "0"))
      .join("");

    return new Promise((resolve) => {
      img.onload = () => {
        canvas.width = img.width;
        canvas.height = img.height;
        ctx.drawImage(img, 0, 0);
        const imageData = ctx.getImageData(0, 0, img.width, img.height);
        const data = imageData.data;

        // Embed binary data into alpha channel
        for (let i = 0; i < binaryData.length; i++) {
          data[i * 4 + 3] = (data[i * 4 + 3] & ~1) | parseInt(binaryData[i], 2); // Alpha channel
        }

        ctx.putImageData(imageData, 0, 0);
        resolve(canvas.toDataURL());
      };
      img.src = URL.createObjectURL(imageFile);
    });
  };

  // Get the modified image as Base64
  const modifiedImageBase64 = await embedJsonInImage(imageFile, jsonObject);

  // Insert image into Word document
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertInlinePictureFromBase64(modifiedImageBase64, Word.InsertLocation.end);
    await context.sync();
  });
}
