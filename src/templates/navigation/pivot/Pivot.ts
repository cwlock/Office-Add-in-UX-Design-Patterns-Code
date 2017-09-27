// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See LICENSE in the project root for license information.

declare var fabric: any;

(($) => {
    $(document).ready(() => {
        Office.initialize = () => { 
            initializePivot(); 
        };

        function initializePivot() {           
            $('.ms-Pivot').map((i, pivot)=>{ new fabric['Pivot'](pivot)});
        }
    });
})(jQuery);