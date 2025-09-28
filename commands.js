let currentSendEvent = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('SafeDocs MVP: Complemento cargado correctamente');
    
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    
    console.log('SafeDocs MVP: Handler de envío registrado');
  }
});

function onMessageSendHandler(event) {
  console.log('SafeDocs MVP: ¡Envío interceptado!');
  
  currentSendEvent = event;
  
  try {
    showConfirmationDialog();
  } catch (error) {
    console.error('SafeDocs MVP: Error al mostrar diálogo:', error);
    event.completed({ allowSend: false });
  }
}

function showConfirmationDialog() {
  const dialogUrl = 'https://andressalazarferro-bit.github.io/safedocs1-mvp/dialog.html';
  
  const dialogOptions = {
    height: 300,
    width: 400,
    displayInIframe: true
  };
  
  console.log('SafeDocs MVP: Mostrando diálogo de confirmación...');
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    dialogOptions,
    function(dialogResult) {
      if (dialogResult.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = dialogResult.value;
        
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(messageEvent) {
          handleUserResponse(messageEvent, dialog);
        });
        
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function(eventArgs) {
          console.log('SafeDocs MVP: Evento de diálogo:', eventArgs.type);
          
          if (eventArgs.type === Office.EventType.DialogClosed) {
            console.log('SafeDocs MVP: Diálogo cerrado por el usuario - Bloqueando envío');
            if (currentSendEvent) {
              currentSendEvent.completed({ allowSend: false });
              currentSendEvent = null;
            }
          }
        });
        
      } else {
        console.error('SafeDocs MVP: Error al abrir diálogo:', dialogResult.error);
        
        if (currentSendEvent) {
          currentSendEvent.completed({ allowSend: false });
          currentSendEvent = null;
        }
      }
    }
  );
}

function handleUserResponse(messageEvent, dialog) {
  try {
    const userResponse = JSON.parse(messageEvent.message);
    
    console.log('SafeDocs MVP: Respuesta del usuario:', userResponse);
    
    dialog.close();
    
    if (userResponse.action === 'send') {
      console.log('SafeDocs MVP: Usuario confirma envío - Permitiendo envío');
      
      if (currentSendEvent) {
        currentSendEvent.completed({ allowSend: true });
      }
      
    } else if (userResponse.action === 'cancel') {
      console.log('SafeDocs MVP: Usuario cancela envío - Bloqueando envío');
      
      if (currentSendEvent) {
        currentSendEvent.completed({ allowSend: false });
      }
      
    } else {
      console.error('SafeDocs MVP: Respuesta inválida del usuario:', userResponse);
      
      if (currentSendEvent) {
        currentSendEvent.completed({ allowSend: false });
      }
    }
    
  } catch (error) {
    console.error('SafeDocs MVP: Error al procesar respuesta:', error);
    
    if (currentSendEvent) {
      currentSendEvent.completed({ allowSend: false });
    }
  } finally {
    currentSendEvent = null;
  }
}
