console.log(Math.random());


(function(){
	let imgs = document.getElementsByTagName('source');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].srcset = "";
	}	
})();

(function(){
	let imgs = document.getElementsByTagName('img');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].src = ""
		imgs[i].srcset = ""
	}
})();


