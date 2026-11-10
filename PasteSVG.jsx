// ============================================
// PasteSVG.jsx - سكريبت لصق SVG في InDesign
// الاستخدام: انسخ كود SVG من Flaticon
// ثم شغّل هذا السكريبت في InDesign
// ============================================

(function () {

    // --- 1. اقرأ كود SVG من الكليبورد ---
    var svgCode = "";

    try {
        // طريقة Mac: استخدم osascript لقراءة الكليبورد
        var clipApp = new File("/tmp/read_clip.sh");
        clipApp.open("w");
        clipApp.write("#!/bin/bash\npbpaste");
        clipApp.close();

        // شغّل الأمر
        app.doScript(
            'do shell script "chmod +x /tmp/read_clip.sh && /tmp/read_clip.sh > /tmp/svg_clip_output.txt"',
            ScriptLanguage.APPLESCRIPT_LANGUAGE
        );

        // اقرأ الناتج
        var outFile = new File("/tmp/svg_clip_output.txt");
        if (outFile.exists) {
            outFile.open("r");
            svgCode = outFile.read();
            outFile.close();
        }
    } catch (e) {
        // Fallback: اطلب من المستخدم يلصق الكود يدوياً
        svgCode = "";
    }

    // --- تحقق إن الكود فيه SVG ---
    if (!svgCode || svgCode.indexOf("<svg") === -1) {
        // اطلب من المستخدم إدخال الكود يدوياً
        svgCode = prompt(
            "الكليبورد ما فيه كود SVG.\nالصق كود SVG هنا يدوياً:",
            "<svg xmlns='http://www.w3.org/2000/svg'></svg>"
        );
        if (!svgCode || svgCode.indexOf("<svg") === -1) {
            alert("❌ ما في كود SVG صالح. تأكد من نسخ الكود من Flaticon أولاً.");
            return;
        }
    }

    // --- 2. احفظ SVG في ملف مؤقت ---
    var tempSVG = new File("/tmp/indesign_paste_svg.svg");
    tempSVG.open("w");
    tempSVG.write(svgCode);
    tempSVG.close();

    if (!tempSVG.exists) {
        alert("❌ فشل في حفظ ملف SVG المؤقت.");
        return;
    }

    // --- 3. تأكد إن في مستند مفتوح ---
    if (app.documents.length === 0) {
        alert("❌ ما في مستند InDesign مفتوح.");
        return;
    }

    var doc = app.activeDocument;
    var page = doc.pages[0];

    // --- 4. اعرف حجم الصفحة ---
    var pageWidth  = doc.documentPreferences.pageWidth;
    var pageHeight = doc.documentPreferences.pageHeight;

    // حجم الأيقونة (يمكن تعديله)
    var iconSize = 50; // بالـ points (حوالي 1.7 سم)

    // ضع الأيقونة في المنتصف
    var x = (pageWidth  - iconSize) / 2;
    var y = (pageHeight - iconSize) / 2;

    // --- 5. أنشئ frame وضع الـ SVG فيه ---
    try {
        var frame = page.rectangles.add({
            geometricBounds: [y, x, y + iconSize, x + iconSize],
            strokeWeight: 0,
            strokeColor: doc.swatches.itemByName("None")
        });

        frame.place(tempSVG);
        frame.fit(FitOptions.PROPORTIONALLY);
        frame.fit(FitOptions.FRAME_TO_CONTENT);

        // حدد العنصر
        app.select(frame);

        alert("✅ تم إضافة الأيقونة SVG بنجاح!\n\nيمكنك الآن تحريكها وتغيير حجمها.");

    } catch (e) {
        alert("❌ خطأ في إضافة الـ SVG:\n" + e.message +
              "\n\nتأكد إن كود SVG صحيح.");
    }

})();
