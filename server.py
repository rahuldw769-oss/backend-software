
        if not response.ok:

            raise HTTPException(

                status_code=500,

                detail=(
                    "Ollama error: "
                    + response.text
                )

            )


        result = response.json()


        answer = ""


        if isinstance(
            result,
            dict
        ):

            message_data = result.get(
                "message"
            )


            if isinstance(
                message_data,
                dict
            ):

                answer = text(
                    message_data.get(
                        "content"
                    )
                )


            if not answer:

                answer = text(
                    result.get(
                        "response"
                    )
                )


        if not answer:

            answer = (
                "AI ne koi response "
                "generate nahi kiya."
            )


        return {

            "status":
                "success",

            "answer":
                answer

        }


    except HTTPException:

        raise


    except requests.exceptions.ConnectionError:

        raise HTTPException(

            status_code=503,

            detail=(
                "Local AI server (Ollama) "
                "connect nahi ho raha. "
                "Check karo ki Ollama running hai "
                "aur OLLAMA_URL sahi hai."
            )

        )


    except requests.exceptions.Timeout:

        raise HTTPException(

            status_code=504,

            detail=(
                "Local AI response mein "
                "bahut time lag raha hai. "
                "Ollama model ya computer load check karo."
            )

        )


    except Exception as e:

        raise HTTPException(

            status_code=500,

            detail="AI error: " + str(e)

        )
