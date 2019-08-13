<?php

$test = new sharepoint();
var_dump( $test->token);


class sharepoint {
  // variables for sharepoint
  private $grant_type;
  private $client_id;
  private $client_secret;
  private $scope;
  private $directory_id;
  private $url;
  private $host_name;
  public $token;
  
  public function __construct() {
    
    // set credentials
    $this->grant_type       = 'client_credentials';
    $this->client_id        = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
    $this->client_secret    = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
    $this->scope            = 'https://graph.microsoft.com/.default';
    $this->directory_id     = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
    $this->host_name        = 'yourcompany.sharepoint.com';
    $this->url['token']     = 'https://login.microsoftonline.com/'.$this->directory_id.'/oauth2/v2.0/token';
    $this->url['sites']     = 'https://graph.microsoft.com/v1.0/sites/';
    
    // get token for operations
    $this->token            = $this->Token_get();
  }
  
  public function Token_get() {
    $curl = curl_init();
    curl_setopt_array($curl, array(
      CURLOPT_URL => $this->url['token'],
      CURLOPT_RETURNTRANSFER => true,
      CURLOPT_ENCODING => "",
      CURLOPT_MAXREDIRS => 10,
      CURLOPT_TIMEOUT => 30,
      CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
      CURLOPT_CUSTOMREQUEST => "POST",
      CURLOPT_POSTFIELDS => "grant_type=$this->grant_type&client_id=$this->client_id&client_secret=$this->client_secret&scope=$this->scope",
      CURLOPT_HTTPHEADER => array(
        "Accept: */*",
        "Accept-Encoding: gzip, deflate",
        "Cache-Control: no-cache",
        "Connection: keep-alive",
      ),
    ));
    $response = curl_exec($curl);
    $err = curl_error($curl);
    $errNo = curl_errno($curl);
    curl_close($curl);
    if ($err) {
      return array ('error' => $err, 'error_number' => $errNo);
    } else {
      $array = json_decode($response, true);
      $array['expires_timestamp'] = time() + 3540; // sets expire-timestamp to 59 minutes in the future
      return $array;
    }
  }

  public function SiteId_get($relativePath) {
    $url          = $this->url['sites'].$this->host_name .':/'.$relativePath;
    $postfields   = '';
    $method       = 'GET';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method);
    return $return;
  }
  
  public function ListId_get($siteId) {
    $url          = $this->url['sites'].$siteId.'/lists';
    $postfields   = '';
    $method       = 'GET';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method);
    return $return;
  }
  
  public function ListItems_get($siteId, $listId, $query='') {
    $url          = $this->url['sites'].$siteId.'/lists/'.$listId.'/items'.$query;
    $postfields   = '';
    $method       = 'GET';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method);
    return $return;
  }
  
  private function curl($url, $postfields, $method='GET') {
    $curl = curl_init();
    curl_setopt_array($curl, array(
      CURLOPT_URL => $url,
      CURLOPT_RETURNTRANSFER => true,
      CURLOPT_ENCODING => "",
      CURLOPT_MAXREDIRS => 10,
      CURLOPT_TIMEOUT => 30,
      CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
      CURLOPT_CUSTOMREQUEST => $method,
      CURLOPT_POSTFIELDS => $postfields,
      CURLOPT_HTTPHEADER => array(
        "Accept: */*",
        "Accept-Encoding: gzip, deflate",
        "Cache-Control: no-cache",
        "Connection: keep-alive",
        "Authorization:Bearer " . $this->token['access_token']
      ),
    ));
    $response = curl_exec($curl);
    $err = curl_error($curl);
    $errNo = curl_errno($curl);
    curl_close($curl);
    if ($err) {
      return array ('error' => $err, 'error_number' => $errNo);
    } else {
      return json_decode($response, true);
    }
  }
  
}

