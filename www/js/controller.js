angular.module('starter.controllers', ['wj'])

.controller('DashCtrl', function($scope, $state, $ionicPopup, $ionicListDelegate) {    
    $scope.tasks =
      [
        {title: "First", time: 1, completed: true},
        {title: "Second",time: 2,completed: false},
        {title: "Third", time: 5, completed: false},
      ];

      $scope.newTask = function() {
        $state.go("test");
       /* $ionicPopup.prompt({
          title : "New Task",
          template: "Enter task:",
          inputPlaceholder: "What do you need to do?",
          okText: 'Create task'
        }).then(function(res) {    // promise 
          if (res) $scope.tasks.push({title: res, completed: false});
        })*/
      };
    
     $scope.edit = function(task) {
      $scope.data = { response: task.title };
      $ionicPopup.prompt({
        title: "Edit Task",
        scope: $scope
      }).then(function(res) {    // promise 
        if (res !== undefined) task.title = $scope.data.response;
        $ionicListDelegate.closeOptionButtons()
     })
    };
})

.controller('ChatsCtrl', function($scope) {})

.controller('TestCtrl', function($scope, $ionicHistory) {

   $scope.someValue = 3.5; 
   $scope.goBack = function() {     
      $ionicHistory.goBack();
   };
})


