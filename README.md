# Rà soát KHCN — Vercel deployment

## Deploy with Vercel Drop

1. Sign in at [Vercel Drop](https://vercel.com/drop).
2. Drop `khcn-vercel.zip` on the page, choose a project name and select **Deploy**. If uploading an extracted folder instead, select the one containing `index.html` and `vercel.json`.

## Deploy from Git or the Vercel CLI

Extract the ZIP. Import its contents as a project, with the folder containing `vercel.json` as the project root. For the CLI, run `vercel --prod` from that folder. The included configuration selects **Other**, serves the project root, and skips both installation and build. No environment variables are required.

## Use the app

1. Select **Chọn mẫu Word** and choose the mapped `.docx` template from your device.
2. Select the mapped `.xlsx`, review the customer sheets, and choose **Tạo bộ báo cáo**.
3. Download the Word reports and the workbook with `TONG_HOP` as the last sheet, individually or as a ZIP.

Keep the Word template and customer workbooks on your device. The app reads them in the browser. To preview locally, open `index.html` from the extracted folder and leave its asset files in place.

[Vercel Drop documentation](https://vercel.com/docs/drop) · [Build settings](https://vercel.com/docs/builds/configure-a-build)
