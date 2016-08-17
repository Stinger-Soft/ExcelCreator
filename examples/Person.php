<?php
class Person {

	protected $name;

	protected $email;

	public function __construct($name, $email) {
		$this->name = $name;
		$this->email = $email;
	}

	public function getName() {
		return $this->name;
	}

	public function setName($name) {
		$this->name = $name;
		return $this;
	}

	public function getEmail() {
		return $this->email;
	}

	public function setEmail($email) {
		$this->email = $email;
		return $this;
	}
}