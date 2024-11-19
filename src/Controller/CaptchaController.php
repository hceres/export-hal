<?php


namespace App\Controller;


use App\Service\ApiPiste;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;


/**
 * Class WsController
 * @package App\Controller
 */

class CaptchaController extends AbstractController
{

    /**
     * @Route("/captcha/{endpoint}", methods={"GET"}, name="captcha" )
     * @param Request $request
     * @param ApiPiste $apiService
     * @param string $endpoint
     * @return Response
     * @throws \Exception
     */

    public function captchaRequest(Request $request, ApiPiste $apiService, string $endpoint){
        $endpoint .= "?".$request->getQueryString();
        return new Response($apiService->getContentResponse($endpoint));
    }

    /**
     * @Route("/captcha/validation", methods={"POST"}, name="captcha-validation")
     * @param Request $request
     * @param ApiPiste $apiService
     * @return Response
     * @throws \Exception
     */

    public function captchaValidation(Request $request, ApiPiste $apiService) {
        return new Response($apiService->isCaptchaValid($request->getContent()));
    }

}
