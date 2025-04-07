Office.onReady(() => {
  // Register command handlers
  Office.actions.associate("COMMAND_ID", function() {
    // Command handler implementation
    console.log("Command executed");
  });
}); 