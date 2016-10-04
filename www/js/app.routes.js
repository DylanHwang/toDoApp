(function () {

	angular
		.module('starter')
		.config(function ($stateProvider, $urlRouterProvider) {

              $stateProvider

               .state('tab', {
                url: '/tab',
                abstract: true,
                templateUrl: 'templates/tabs.html'
                })                
                // Each tab has its own nav history stack:

                .state('tab.toDoList', {
                    url: '/toDoList',
                    views: {
                    'tab-toDoList': {
                        templateUrl: 'templates/tab-toDoList.html',
                        controller: 'DashCtrl'
                    }
                    }
                })

                .state('tab.chats', {
                    url: '/chats',
                    views: {
                        'tab-chats': {
                        templateUrl: 'templates/tab-chats.html',
                        controller: 'ChatsCtrl'
                        }
                    }
                 })                

                .state('tab.account', {
                    url: '/account',
                    views: {
                    'tab-account': {
                        templateUrl: 'templates/tab-account.html',
                        controller: 'AccountCtrl'
                    }
                    }
                })
                
                .state('tab.newTask', {
                    url: '/newTask',    
                    views: {
                         'tab-toDoList': {
                          templateUrl: 'templates/newTask.html',
                    controller: 'TestCtrl'
                    }                  
                    }        
                });

                $urlRouterProvider.otherwise('/tab/toDoList');

});

})();