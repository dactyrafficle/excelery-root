  let changeColorButton = document.getElementById('changeColorButton');

  chrome.storage.sync.get('color', function(data) {
    changeColorButton.style.backgroundColor = data.color;
    changeColorButton.setAttribute('value', data.color);
  });
	
// now we have to add the logic that makes the button actually do something

  changeColorButton.onclick = function(element) {
    //console.log('ok you clicked ' + Math.floor(Math.random()*899+100));
    let color = element.target.value;
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      chrome.tabs.executeScript(
          tabs[0].id, // you must indicate the tab being accessed to get to dom
					{file: "main.js"}); // this works, but you must include the code file in the manifest
					
					// why is it written like this? and not just outside?
          //{code: 'console.log("hi.");document.body.style.backgroundColor = "' + color + '";'});
						//{code: y});
							//{code: y});
    });
		
		// why not just like this? not work. it not work.
		/*
		(function () {
			var script = document.createElement("script");
			script.src ="code.js";
			document.body.appendChild(script);
		})();
		*/
		
		
  };
	
	// so console.log("hi"); gets printed to the proper console on the tab - that is exciting!
	
	changeColorButton.addEventListener('click', function(e) {
		console.log('ok you clicked ' + Math.floor(Math.random()*899+100)); // works, in the outer console.
		
		
		
	});
	
	// interesting
	var x = '(function(){alert("hi!")})();' // can do like this also.
	
	var y = '(function () {var script = document.createElement("script"); script.src ="code.js";document.body.appendChild(script);})();'