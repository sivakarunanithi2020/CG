import React from "react";

export default function App() {
  
  const handleSetCategory = () => {
    let categoriesToAdd = ["HelloSF"];
    Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to add category: " + JSON.stringify(asyncResult.error));
      } else {
        console.log(`Category "${categoriesToAdd}" successfully added to the item.`);
      }
    });
  };

  return (
    <div>
      <button onClick={handleSetCategory}>Set Category</button>
    </div>
  );
}
