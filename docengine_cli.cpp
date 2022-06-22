/**
  Copyright (c) 2016-2022, Smart Engines Service LLC
  All rights reserved.
*/

#include <cstring>
#include <string>
#include <memory>
#include <cstdio>
#include <regex>

#ifdef _MSC_VER
#pragma warning( disable : 4290 )
#include <windows.h>
#endif // _MSC_VER

#include <secommon/se_common.h>

#include <docengine/doc_engine.h>
#include <docengine/doc_result.h>
#include <docengine/doc_session_settings.h>
#include <docengine/doc_session.h>
#include <docengine/doc_processing_settings.h>
#include <docengine/doc_graphical_structure.h>
#include <docengine/doc_views_collection.h>
#include <docengine/doc_document.h>
#include <docengine/doc_fields.h>

// Here we simply output the recognized fields in a JSON format
void OutputRecognitionResult(
    const se::doc::DocResult& recog_result) {

  if (recog_result.GetDocumentsCount() == 0) {
    printf("{}\n");
  } else {
    const se::doc::Document& doc = recog_result.DocumentsBegin().GetDocument();
    printf("{\"DOCTYPE\": \"%s\"", doc.GetAttribute("type"));
    for (auto f_it = doc.TextFieldsBegin();
         f_it != doc.TextFieldsEnd();
         ++f_it) {
      std::string escaped_value = std::regex_replace(
          f_it.GetField().GetOcrString().GetFirstString().GetCStr(), 
          std::regex("\""), "\\\"");
      printf(",\"%s\": \"%s\"", 
          f_it.GetKey(),
          escaped_value.c_str());
    }
    printf("}\n");
  }
}


int main(int argc, char **argv) {
#ifdef _MSC_VER
  SetConsoleOutputCP(65001);
#endif // _MSC_VER

  // 1st argument - path to the image to be recognized
  // 2nd argument - path to the configuration bundle
  // 3rd argument - document types mask, "*" by default
  if (argc != 3 && argc != 4) {
    printf("Version %s. Usage: "
           "%s <image_path> <bundle_zip_path> [document_types]\n",
           se::doc::DocEngine::GetVersion(), argv[0]);
    return -1;
  }

  const std::string image_path = argv[1];
  const std::string config_path = argv[2];
  const std::string document_types = (argc >= 4 ? argv[3] : "*");

  try {
    // Creating the recognition engine object - initializes all internal
    //     configuration structure. Second parameter to the factory method
    //     is the lazy initialization flag (true by default). If set to
    //     false, all internal objects will be initialized here, instead of
    //     waiting until some session needs them.
    std::unique_ptr<se::doc::DocEngine> engine(
        se::doc::DocEngine::Create(config_path.c_str(), true));

    // Before creating the session we need to have a session settings
    //     object. Such object can be created only by acquiring a
    //     default session settings object from the configured engine.
    std::unique_ptr<se::doc::DocSessionSettings> session_settings(
        engine->CreateSessionSettings());

    // Setting mode to "universal", assuming configuration bundle "bundle_docengie_photo.se"
    session_settings->SetCurrentMode("universal");
    // For starting the session we need to set up the mask of document types
    //     which will be recognized.
    session_settings->AddEnabledDocumentTypes(document_types.c_str());

    // Creating a session object - a main handle for performing recognition.
    std::unique_ptr<se::doc::DocSession> session(
        engine->SpawnSession(*session_settings, ${put_yor_personalized_signature_from_doc_README.html}));

    // Creating default image processing settings - needs to be created before
    //     passing an image for processing. This specifies how the session
    //     should process the updated source.
    std::unique_ptr<se::doc::DocProcessingSettings> proc_settings(
        session->CreateProcessingSettings());

    // Creating an image object which will be used as an input for the session
    std::unique_ptr<se::common::Image> image(
        se::common::Image::FromFile(image_path.c_str()));

    // Registering input image in the session,
    //     setting it up as the current source and processing the image
    int image_id = session->RegisterImage(*image);
    proc_settings->SetCurrentSourceID(image_id);
    session->Process(*proc_settings);

    // Obtaining the recognition result
    const se::doc::DocResult& result = session->GetCurrentResult();

    // Printing the contents of the recognition result
    OutputRecognitionResult(result);

  } catch (const se::common::BaseException& e) {
    printf("Exception thrown: %s\n", e.what());
    return -1;
  }

  return 0;
}
