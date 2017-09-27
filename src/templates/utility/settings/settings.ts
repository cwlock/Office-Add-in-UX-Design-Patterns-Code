// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See LICENSE in the project root for license information.

/**
 * Settings Template
 *
 * A basic add-in layout for an drill down settings page.
 * Includes the back button navigation.
 *
 */

declare var fabric: any;

(($) => {
    $(document).ready(() => {
        Office.initialize = () => { 
            initializeDropdown(); 
            initializeToggle();
            initializeButton();
        };

        function initializeDropdown() {           
            $('.ms-Dropdown').map((i, dropdown)=>{ new fabric['Dropdown'](dropdown)});
        }

        function initializeToggle() {            
            $('.ms-Toggle').map((i, toggle)=>{ new fabric['Toggle'](toggle)});
        }

        function initializeButton() {           
            $('.ms-Button').map((i, button)=>{ new fabric['Button'](button)});
        }
    });
})(jQuery);
