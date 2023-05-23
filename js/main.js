Office.onReady();

async function addSpace() {
     await PowerPoint.run(async (context) => {
          let slides = context.presentation.getSelectedSlides();
          slides.load("items");
          await context.sync();
          let curSlideID = slides.items[0].id;
          context.presentation.load("slides");
          await context.sync();
          let activeSlide = context.presentation.slides.getItem(curSlideID);
          activeSlide.load("shapes");
          await context.sync();
          let shapes = activeSlide.shapes;
          shapes.load("items");
          await context.sync();
          for (let i = 0; i < shapes.items.length; i++) {
               let tmpObj = new Object();
               tmpObj.slideID = curSlideID;
               tmpObj.shapeID = shapes.items[i].id;
               let curRange = shapes.items[i].textFrame.textRange;
               shapes.items[i].textFrame.textRange.load("text");
               await context.sync();
               let curTextContent = shapes.items[i].textFrame.textRange.text;
               let regex1 = /([\u4E00-\u9FFF])([A-Za-z0-9])/g;
               let regex2 = /([A-Za-z0-9])([\u4E00-\u9FFF])/g;
               await addspace(regex1, 1);
               await addspace(regex2, 1);
               async function addspace(regex, insertIndex) {
                    while ((match = regex.exec(curTextContent)) != null) {
                         let tmpRange = curRange.getSubstring(match.index, 1);
                         tmpRange.load("text");
                         await context.sync();
                         console.log("Matched text content: " + match[0]);
                         console.log("Matched index: " + match.index);
                         console.log("original texRange: " + tmpRange.text);
                         tmpRange.text = match[1] + " ";
                         console.log("added texRange: " + tmpRange.text);
                         curTextContent = curTextContent.slice(0, match.index + insertIndex) + " " + curTextContent.slice(match.index + insertIndex);
                         console.log("updated total text rage: " + curTextContent);
                         console.log("===========================");
                    }
                    return context.sync();
               }
          }
     });
}
