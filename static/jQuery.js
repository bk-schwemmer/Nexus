function confirm_overwrite() {
            overwrite = confirm('Are you sure you wish to overwrite the temp file?');
            if (overwrite) {
                // ajax function to overwrite temp file in background without refreshing webpage
                $(function() {
                  $('a#overwrite_button').on('click', function(e) {
                    e.preventDefault()
                    $.getJSON('/background_overwrite',
                        function(data) {
                      // $('#overW').show();
                    });
                    return false;
                  });
                });
            } else {
                return;
            }
        }

/*fetch('/log')
    .then(function (response) {
        return response.json();
    }).then(function (text) {
        console.log('GET response:');
        console.log(text.no_new_hires);
});*/
