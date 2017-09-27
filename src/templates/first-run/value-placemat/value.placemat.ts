// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See LICENSE in the project root for license information.

declare var fabric: any;

(($) => {
    $(document).ready(() => {
        Office.initialize = () => { 
            initializeButton();
        };

        function initializeButton() {           
            $('.ms-Button').map((i, button)=>{ new fabric['Button'](button)});
        }
    });
})(jQuery);