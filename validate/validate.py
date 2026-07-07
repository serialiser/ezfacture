import saxonche
from lxml import etree


def validate_xml(xml_file, xslt_file):
    with saxonche.PySaxonProcessor(license=False) as proc:

        xslt_processor = proc.new_xslt30_processor()
        xslt_executable = xslt_processor.compile_stylesheet(stylesheet_file=xslt_file)

        result = xslt_executable.transform_to_string(source_file=xml_file)

        # Parse the result string into an XML tree
        try:
            result_tree = etree.fromstring(result.encode('utf-8'))
        except etree.XMLSyntaxError as e:
            print(f"Erreur lors de l'analyse du résultat XML : {str(e)}")
            return

        # Define the SVRL namespace
        svrl_ns = {"svrl": "http://purl.oclc.org/dsdl/svrl"}

        # Check for validation errors
        failed_asserts = result_tree.xpath("//svrl:failed-assert", namespaces=svrl_ns)

        if failed_asserts:
            print("Le fichier XML n'est pas valide:")
            for error in failed_asserts:
                location = error.get("location", "Emplacement inconnu")
                message = error.find("svrl:text", namespaces=svrl_ns)
                message_text = message.text if message is not None else "Message non disponible"
                print(f"location : {location} :\n -> Description : {message_text} \n\n")
        else:
            print("Le fichier XML est valide.")


xml_path = "example_invoice.xml"
xslt_path = "EN16931-UBL-validation.xslt"

# Validate the XML document
validate_xml(xml_path, xslt_path)
