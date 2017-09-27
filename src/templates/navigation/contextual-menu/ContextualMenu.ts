// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See LICENSE in the project root for license information.

declare var fabric: any;

(($) => {
    $(document).ready(() => {
        Office.initialize = () => { 
            initializeContextualMenu(); 
        };

        function initializeContextualMenu() {
            let $button = $('#mainMenu');            
            $('.ms-ContextualMenu').map((i, menu)=>{ new fabric['ContextualMenu'](menu, $button[0])});
        }
    });
})(jQuery);