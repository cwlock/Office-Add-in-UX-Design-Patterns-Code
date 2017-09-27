/// // Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.

declare var fabric: any;

(($) => {
    "use strict";

    let app = (<any>window).app || {};
    app.firstrun = {};
    app.firstrun.stage = 1;

    
    $(document).ready(function () {
        Office.initialize = () => { 
            initializeButton();
        }

        // Initializes the button component
        function initializeButton() {           
            $('.ms-Button').map((i, button)=>{ new fabric['Button'](button)});
        }

        // $('.dp-carousel').hover(()=>{
        //     $('.changePage').fadeIn(300);
        // }, () => {
        //     $('.changePage').fadeOut(300);
        // })

        app.firstrun.totalPages = $('#pageMarkers').get(0).childElementCount;

        // Navigates to a different stage
        app.firstrun.newStage = function () {
            if (this.id === 'next') {
                app.firstrun.nextStage();
            } else {
                app.firstrun.previousStage();
            }
        };

        $('.changePage').click(app.firstrun.newStage);

        // Navigates to the next stage of the First Run experience.
        app.firstrun.nextStage = function () {
            var nextToLastPage = app.firstrun.totalPages - 1;
            switch (app.firstrun.stage) {
                case nextToLastPage:
                    // Render the special last page UI, e.g., the Next button disappears.
                    app.firstrun.showLastStage();
                    break;
                default:
                    // Note the "+ 1" in the parameter, so the next page is rendered.
                    app.firstrun.showIntermediateStage(app.firstrun.stage + 1);
                    break;
            }
        };

        // Navigates to the previous stage of the First Run experience.
        app.firstrun.previousStage = function () {

            switch (app.firstrun.stage) {
                case 2:
                    // Render the special first page UI, e.g., the Previous button disappears.
                    app.firstrun.showStartStage();
                    break;
                default:
                    // Note the "- 1" in the parameter, so the previous page is rendered.
                    app.firstrun.showIntermediateStage(app.firstrun.stage - 1);
                    break;
            }
        };

        app.firstrun.showStartStage = function() {

            // The UI pattern that is unique to the start stage
            $('#skip').attr('class', 'ms-font-m');
            $('#dp-slide-1').attr('class', 'dp-carousel--content__slide');
            $('#prev').attr('class', 'displayNone');
            $('#dp-carousel--content__button').attr('class', 'displayNone');
            // UI changes that apply to all stage transitions
            app.firstrun.setStageUI(1);
        };
       
        app.firstrun.showIntermediateStage = function (nextStage:any) {

            // The UI pattern that is unique to intermediate stages
            $('#skip').attr('class', 'ms-font-m');
            $('#next').attr('class', 'change-stage');
            $('#prev').attr('class', 'change-stage');
            $('#dp-carousel--content__button').attr('class', 'displayNone');
            // UI changes that apply to all stage transitions
            app.firstrun.setStageUI(nextStage);
        };

        app.firstrun.showLastStage = function () {

            // The UI pattern that is unique to the last stage
            $('#skip').attr('class', 'hide ms-font-m');
            $('#next').attr('class', 'displayNone');
            $('#previous').attr('class', 'change-stage ms-font-m ms-fontColor-white');
            $('#dp-carousel--content__button').attr('class', 'ms-Button ms-Button--primary');
            // UI changes that apply to all stage transitions
            app.firstrun.setStageUI(app.firstrun.totalPages);
        };

        app.firstrun.setStageUI = function (stage:any) {
            $('#dp-slide-' + stage).attr('class', 'dp-carousel--content__slide');
            $('#dp-slide-' + stage).show().siblings('div').hide();
            app.firstrun.setStageProgress(stage);
            app.firstrun.stage = stage;
        }

        app.firstrun.setStageProgress = function (stage:any) {
            $('#dp-dot-' + stage).attr('class', 'dot');
            $('#dp-dot-' + stage).siblings().attr('class', 'dot ms-fontColor-neutralTertiaryAlt');
        }

        app.firstrun.showStartStage();
    });    
    (<any>window).app = app;
})(jQuery);