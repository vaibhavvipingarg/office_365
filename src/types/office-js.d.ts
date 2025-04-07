// This file provides simple type declarations for Office.js

declare namespace Office {
  function onReady(callback: () => void): void;
  
  enum AsyncResultStatus {
    Succeeded = "succeeded",
    Failed = "failed"
  }

  enum CoercionType {
    Text = "text",
    Html = "html",
    Image = "image"
  }

  interface AsyncResult {
    status: AsyncResultStatus;
    error: any;
    value: any;
  }

  namespace actions {
    function associate(commandId: string, handler: () => void): void;
  }

  namespace context {
    const document: {
      setSelectedDataAsync(
        data: any,
        options: { coercionType: CoercionType },
        callback: (result: AsyncResult) => void
      ): void;
    };
  }
} 