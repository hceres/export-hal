<?php

namespace App\Controller;

use App\Service\ExportHal;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\Session\Session;
use Symfony\Component\Routing\Annotation\Route;


class ExportHalController extends AbstractController
{
    private string $captchaActif;

    /**
     * ExportHalController constructor.
     * @param string $halCaptchaActif
     */
    public function __construct(string $captchaActif)
    {
        $this->captchaActif = $captchaActif;
    }

    /**
     * @Route("/")
     * @return Response
     */
    public function index():Response
    {
        return $this->redirectToRoute('exportHal');
    }

    /**
     * @Route("/hal", name="exportHal",  methods={"GET"})
     * @return Response
     */
    public function exporthal()
    {
        return $this->render('core/exporthal.html.twig');
    }

    /**
     * @Route("/exporthal", name="exportHalSubmit", methods={"GET", "POST"})
     * @param Request $request
     * @return JsonResponse
     * @throws \Exception
     */
    public function exporthalSubmit(Request $request)
    {
        $exportHal = new ExportHal();
        $result = $exportHal->getResult(
            $request->get('recherche'),
            $request->get('idshal'),
            $request->get('idcoll'),
            $request->get('dateDeb'),
            $request->get('dateFin'),
            $request->get('equipelabo'),
        );

        $json = new JsonResponse();
        $json->setContent($result);
        return  $json;
    }
}
