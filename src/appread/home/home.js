(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      onInitializeComplete();
    });
  };

  function onInitializeComplete() {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, {}, function (asyncResult) {
      if (asyncResult.status === 'failed') {
        console.log('Texifiable: Failed to get the body of the item.');
        return;
      }
 
      var body = asyncResult.value;
      $('#content').html(body);
 
      // Since the content has been changed, we need to tell MathJax to
      // look for mathematics in the page again.
      MathJax.Hub.Queue(
        ["Typeset", MathJax.Hub, "content"],
        onTypesetComplete);
    });
  }

  function onTypesetComplete() {
    $('#typeset-message').css('display', 'none');
    $('#content').css('display', 'block');
  }

})();
