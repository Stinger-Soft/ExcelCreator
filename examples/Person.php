<?php

class Person {

	protected $name;

	protected $email;

	protected $address;

	protected $city;

	protected $zipCode;

	protected $phoneNumber;

	protected $birthday;

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

	public function getAddress() {
		return $this->address;
	}

	public function setAddress($address) {
		$this->address = $address;
		return $this;
	}

	public function getCity() {
		return $this->city;
	}

	public function setCity($city) {
		$this->city = $city;
		return $this;
	}

	public function getZipCode() {
		return $this->zipCode;
	}

	public function setZipCode($zipCode) {
		$this->zipCode = $zipCode;
		return $this;
	}

	public function getPhoneNumber() {
		return $this->phoneNumber;
	}

	public function setPhoneNumber($phoneNumber) {
		$this->phoneNumber = $phoneNumber;
		return $this;
	}

	/**
	 * @return mixed
	 */
	public function getBirthday() {
		return $this->birthday;
	}

	/**
	 * @param mixed $birthday
	 * @return Person
	 */
	public function setBirthday($birthday) {
		$this->birthday = $birthday;
		return $this;
	}


}