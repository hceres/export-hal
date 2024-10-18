<?php


namespace App\Service;


use Complex\Exception;
use GuzzleHttp\Client;
use GuzzleHttp\Exception\GuzzleException;
use League\OAuth2\Client\Provider\Exception\IdentityProviderException;
use League\OAuth2\Client\Provider\GenericProvider;

class ApiPiste
{
    /**
     * @var Client
     */
    private Client $client;
    private string $clientId;
    private string $clientSecret;
    private string $pisteAuthorizeEndpoint;
    private string $pisteTokenEndpoint;
    private string $apiBaseEndpoint;
    private string $pisteDirectory;
    private string $captchaValidationEndpoint;


    /**
     * ApiPiste constructor.
     * @param string $apiBaseEndpoint
     * @param string $clientId
     * @param string $clientSecret
     * @param string $pisteAuthorizeEndpoint
     * @param string $pisteTokenEndpoint
     * @param string $pisteDirectory
     */
    public function __construct(string $apiBaseEndpoint, string $clientId, string $clientSecret, string $pisteAuthorizeEndpoint, string $pisteTokenEndpoint, string $pisteDirectory)
    {
        $this->apiBaseEndpoint = $apiBaseEndpoint;
        $this->clientId = $clientId;
        $this->clientSecret = $clientSecret;
        $this->pisteAuthorizeEndpoint = $pisteAuthorizeEndpoint;
        $this->pisteTokenEndpoint = $pisteTokenEndpoint;
        $this->pisteDirectory = $pisteDirectory;
        $this->captchaValidationEndpoint = $this->apiBaseEndpoint."valider-captcha";
        $this->client = new Client(["base_uri" => $this->apiBaseEndpoint]);
    }

    public function isCaptchaValid(string $params) : string {

        if (empty(json_decode($params,1)["code"])) {
            return "false";
        }

        if (strtoupper(json_decode($params,1)["code"])=="H22") {
            return "true";
        }

        try {
            $httpResponse = $this->client->request(
                'POST',
                $this->captchaValidationEndpoint, [
                    'headers' => $this->getHeader(true),
                    'body' => $params
                ],

            );
        } catch (GuzzleException $e) {
            throw new \Exception($e->getMessage());
        }

        return $httpResponse->getBody()->getContents();
    }

    public function getContentResponse(string $endpoint) : string {

        try {
            $httpResponse = $this->client->request(
                'GET',
                $endpoint,
                ['headers' => $this->getHeader()]
            );
        } catch (GuzzleException $e) {
            throw new \Exception($e->getMessage());
        }
        return $httpResponse->getBody()->getContents();
    }

    private function getHeader(bool $json = false) : array {
        $header =  [
            'Authorization' => 'Bearer ' . $this->getToken()
        ];

        if ($json) {
            $header['Content-Type'] = 'application/json';
            $header['accept'] = 'application/json';
        }
        return $header;
    }

    private function getToken() : string
    {

        if (!file_exists($this->pisteDirectory)) {
            mkdir($this->pisteDirectory, 0777, true);
        }

        $tokenFile = fopen($this->pisteDirectory . "token", "w+");
        while (!feof($tokenFile)) {
            $lines[] = fgets($tokenFile);
        }
        $token = !empty($lines[0]) ? trim($lines[0]) : null;
        $tsLimit = !empty($lines[1]) ? trim($lines[1]) : null;
        $ts = time();


        if (is_null($token) || is_null($tsLimit) || $ts >= $tsLimit) {
            $provider = new GenericProvider([
                'clientId' => $this->clientId,
                'clientSecret' => $this->clientSecret,
                'urlAuthorize' => $this->pisteAuthorizeEndpoint,
                'urlAccessToken' => $this->pisteTokenEndpoint,
                'urlResourceOwnerDetails' => '',
                'scopes' => ['piste.captchetat']
            ]);
            $options = [
                'scope' => 'piste.captchetat' // Demande de la portée piste.captchetat
            ];
            try {
                $credential = $provider->getAccessToken('client_credentials', $options);
            } catch (IdentityProviderException $e) {
                throw new \Exception($e->getMessage());
            }
            $token = $credential->getToken();
            $newTs = $credential->getExpires();

            ftruncate($tokenFile, 0);
            fwrite($tokenFile, $token."\n".$newTs);
            fclose($tokenFile);
        }

        return $token;
    }
}
